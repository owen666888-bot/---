# license_manager.py (修正版 v4.1 - 包含 Base64 填充修复)
import os
import json
import base64
import hashlib
import platform
import subprocess
import requests
import certifi
import traceback
import uuid
from datetime import datetime, timedelta, timezone # Keep timedelta for potential future expiry logic
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import ed25519
from cryptography.exceptions import InvalidSignature
# import winreg # No longer needed for trial registry marks

# --- 配置区 ---
SECRET_KEY = b'huiDXUG6NmD45vu6k09naUywpcIV7flY6XPhILyPx2U=' # 必须与 key_generator.py 一致

MASTER_PUBLIC_KEY_PEM = """-----BEGIN PUBLIC KEY-----
MCowBQYDK2VwAyEAC1nOHDdSyxXK7G0/vugz6SBZGnEPzQMmD08mXdXjw2Y=
-----END PUBLIC KEY-----"""
# ^^^ 将 key_generator.py 生成的真实公钥粘贴在此处 ^^^

APP_NAME = "MailRoutePro"
PRODUCT_KEY_PREFIX = "MRP"

def get_storage_path():
    """获取存储路径"""
    if os.name == 'nt':
        return os.path.join(os.environ['APPDATA'], APP_NAME)
    else:
        return os.path.join(os.path.expanduser('~'), f'.{APP_NAME.lower()}')

STORAGE_PATH = get_storage_path()
os.makedirs(STORAGE_PATH, exist_ok=True)
ACTIVATION_FILE = os.path.join(STORAGE_PATH, 'activation.key')
TRIAL_FILE = os.path.join(STORAGE_PATH, 'trial.dat')
# 检查您的 Firebase 项目设置和云函数部署
# 以下是一些可能的替代 URL 格式，您可以尝试
# 1. 直接使用函数名称（如果函数是在默认区域部署的）
CHINA_GATEWAY_URL = "https://firebasfunction-hxqjmgrsmv.cn-hongkong.fcapp.run"
FIREBASE_ACTIVATION_URL = f"{CHINA_GATEWAY_URL}/activate"
FIREBASE_TRIAL_URL = f"{CHINA_GATEWAY_URL}/trial"

# 2. 使用完整路径（如果函数是作为 HTTP 函数部署的）
# FIREBASE_ACTIVATION_URL = "https://[区域]-[项目ID].cloudfunctions.net/activateKey"
# FIREBASE_TRIAL_URL = "https://[区域]-[项目ID].cloudfunctions.net/startTrial"

# 添加通用请求头
COMMON_HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json"
}

# Bilingual error messages mapping
BILINGUAL_ERROR_MESSAGES = {
    # Activation related errors
    "InvalidKey": "激活失败：产品密钥无效或不存在。\nActivation failed: The product key is invalid or does not exist.",
    "KeyAlreadyUsed": "激活失败：此产品密钥已被使用。\nActivation failed: This product key has already been used.",
    "KeyExpired": "激活失败：此产品密钥已过期。\nActivation failed: This product key has expired.",
    "InvalidSignature": "激活失败：产品密钥签名无效。\nActivation failed: Invalid product key signature.",
    "MachineMismatch": "激活失败：此产品密钥已绑定到其他设备。\nActivation failed: This product key is bound to another device.",
    "ServerError": "激活失败：服务器内部错误。\nActivation failed: Internal server error.",
    
    # Trial related errors
    "TrialExpired": "试用期已结束。\nTrial period has expired.",
    "TrialNotAvailable": "无法启动试用期。\nUnable to start trial period.",
    "TrialAlreadyStarted": "试用期已经开始。\nTrial period has already started.",
    "InvalidMachineId": "无效的机器ID。\nInvalid machine ID.",
    
    # Network related errors
    "NetworkError": "网络连接错误。\nNetwork connection error.",
    "TimeoutError": "请求超时。\nRequest timeout.",
    
    # General errors
    "InvalidRequest": "无效的请求。\nInvalid request.",
    "UnknownError": "发生未知错误。\nAn unknown error occurred."
}

def get_bilingual_error_message(error_code: str) -> str:
    """
    Get bilingual error message for the given error code.
    If the error code is not found, return a default bilingual message.
    """
    return BILINGUAL_ERROR_MESSAGES.get(error_code, f"发生未知错误 (Unknown Error): {error_code}")

# PURCHASE_URL_EN/ZH are no longer strictly needed by license_manager if dialog doesn't show them
# but can be kept in code.py if you have other places for them.

# --- 核心功能区 ---

def get_machine_id():
    """
    生成一个更稳定、更难伪造的复合式机器指纹。
    (macOS 优化版，避免使用 subprocess)
    """
    fingerprints = []

    # 1. 操作系统信息 (非常稳定且安全)
    fingerprints.append(platform.system())    # 'Darwin' for macOS
    fingerprints.append(platform.release())   # e.g., '23.4.0'
    fingerprints.append(platform.version())   # 详细版本信息
    fingerprints.append(platform.machine())   # 'arm64' for Apple Silicon, 'x86_64' for Intel

    # 2. CPU 信息 (稳定)
    fingerprints.append(platform.processor()) # 'arm' or 'i386'

    # 3. 主机名 (相对稳定)
    # import socket; fingerprints.append(socket.gethostname())
    # 主机名用户可以改，但也可以作为一个指纹维度

    # 4. MAC 地址 (关键指纹，使用 uuid.getnode() 获取)
    # uuid.getnode() 会尝试多种方法获取 MAC 地址，比 subprocess 更健壮
    # 它返回一个48位的整数，我们将其转换为十六进制字符串
    try:
        mac_int = uuid.getnode()
        if (mac_int >> 40) % 2 == 0: # 检查是否是有效的单播 MAC 地址
            mac_addr = ':'.join(('%012X' % mac_int)[i:i+2] for i in range(0, 12, 2))
            fingerprints.append(mac_addr)
    except Exception as e:
        print(f"无法获取 MAC 地址: {e}")
        # 即使 MAC 地址获取失败，我们仍然可以用其他信息生成一个 ID
        pass
        
    # [备选方案] 如果需要更独特的硬件ID，可以尝试 ioreg，但风险如前所述
    # 如果你坚持要用，必须提供完整路径，并处理好权限问题
    # try:
    #     command = "/usr/sbin/ioreg -rd1 -c IOPlatformExpertDevice"
    #     # ... subprocess logic ...
    # except Exception:
    #     pass

    # 组合所有指纹并进行哈希
    if not fingerprints:
        # 如果所有指纹都获取失败，生成一个基于启动时间的随机ID，但这会导致每次都变
        # 更好的做法是至少有一个稳定的信息源
        return hashlib.sha256(str(datetime.now()).encode()).hexdigest()

    combined_string = ":".join(filter(None, fingerprints))
    hashed_id = hashlib.sha256(combined_string.encode()).hexdigest()
    
    print(f"生成的机器指纹源数据: {combined_string}")
    print(f"最终的哈希机器码: {hashed_id}")
    
    return hashed_id

# --- 激活与验证管理 ---

def _verify_product_key_signature(product_key_string: str):
    """验证产品密钥签名 (已修复 Base64 填充问题)"""
    try:
        if not product_key_string.startswith(PRODUCT_KEY_PREFIX + "-"):
            return None, "无效的产品密钥前缀"
        
        parts = product_key_string.split('.')
        if len(parts) != 2:
            return None, "产品密钥格式错误 (缺少分隔符)"
        
        payload_b64_unpadded = parts[0].split('-', 1)[1]
        signature_b64_unpadded = parts[1]

        # 修复Base64填充问题
        payload_b64 = payload_b64_unpadded + '=' * (-len(payload_b64_unpadded) % 4)
        signature_b64 = signature_b64_unpadded + '=' * (-len(signature_b64_unpadded) % 4)

        payload_bytes = base64.urlsafe_b64decode(payload_b64.encode('utf-8'))
        signature_bytes = base64.urlsafe_b64decode(signature_b64.encode('utf-8'))
        
        # 加载 PEM 格式的公钥对象
        public_key_obj = serialization.load_pem_public_key(MASTER_PUBLIC_KEY_PEM.encode('utf-8'))
        
        # 确保加载的公钥对象是 Ed25519PublicKey 类型
        if not isinstance(public_key_obj, ed25519.Ed25519PublicKey):
             return None, "公钥类型错误或加载失败"

        #直接使用加载的公钥对象进行验证
        public_key_obj.verify(signature_bytes, payload_bytes)
        
        license_details = json.loads(payload_bytes.decode('utf-8'))
        return license_details, None
    except InvalidSignature:
        return None, "产品密钥签名无效"
    except base64.binascii.Error as b64_err: # 更具体地捕获Base64错误
        return None, f"产品密钥Base64解析错误: {str(b64_err)}"
    except Exception as e:
        # import traceback # 取消注释以进行深入调试
        # print(f"Product key parsing internal error: {traceback.format_exc()}")
        return None, f"产品密钥解析错误: {str(e)}"

def verify_activation():
    """验证本地激活文件 (逻辑不变，但不再检查 trial 状态)"""
    if not os.path.exists(ACTIVATION_FILE):
        return None, {"status_code": "NeedsActivation_NoKey"} # 更明确的状态码
    try:
        fernet = Fernet(SECRET_KEY)
        with open(ACTIVATION_FILE, 'rb') as f:
            encrypted_data = f.read()
        decrypted_data = fernet.decrypt(encrypted_data)
        activation_info = json.loads(decrypted_data.decode())

        if activation_info.get("machine_id") == get_machine_id():
            # 未来可以在此检查 activation_info["license_details"].get("expires_days")
            # if "activation_date" in activation_info and \
            #    "expires_days" in activation_info.get("license_details", {}) and \
            #    activation_info["license_details"]["expires_days"] > 0:
            #     activation_date_obj = datetime.fromisoformat(activation_info["activation_date"])
            #     expiry_date = activation_date_obj + timedelta(days=activation_info["license_details"]["expires_days"])
            #     if datetime.now() > expiry_date:
            #         return {"status": "expired", "error_message": "您的授权已过期。"}, {"status_code": "NeedsActivation_Expired"}
            return activation_info, {"status_code": "Activated"}
        else:
            return {"status": "mismatch", "error_message": "此激活密钥已绑定至其他设备。"}, {"status_code": "NeedsActivation_Mismatch"}
    except Exception as e:
        return {"status": "corrupted", "error_message": f"激活文件损坏或无效 ({str(e)})。"}, {"status_code": "NeedsActivation_Corrupted"}

def activate_product(product_key_string: str):
    """
    Activate product with product key using server-side validation.
    Sends product key and machine ID to server, receives and stores support key if valid.
    
    Args:
        product_key_string (str): The product key to activate with
        
    Returns:
        tuple: (success: bool, message: str)
    """
    print("\n" + "="*50)
    print("[ACTIVATION] Starting product activation process")
    print(f"[ACTIVATION] Product key: {product_key_string}")
    print(f"[ACTIVATION] Using activation URL: {FIREBASE_ACTIVATION_URL}")
    print("="*50 + "\n")
    try:
        # 准备激活请求数据
        activation_payload = {
            "productKey": product_key_string,
            "machineId": get_machine_id()
        }
        
        # 发送激活请求到服务器
        print(f"[ACTIVATION] Sending activation request for machine: {activation_payload['machineId']}")
        print(f"[ACTIVATION] Sending request to: {FIREBASE_ACTIVATION_URL}")
        print(f"[ACTIVATION] Payload: {activation_payload}")
        try:
            print(f"[ACTIVATION] Sending request to server...")
            print(f"[ACTIVATION] Request payload: {json.dumps(activation_payload, indent=2)}")
            print(f"[ACTIVATION] Request headers: {json.dumps({'Content-Type': 'application/json'}, indent=2)}")
            
            response = requests.post(
                FIREBASE_ACTIVATION_URL,
                json=activation_payload,
                headers=COMMON_HEADERS,
                timeout=30,
                verify=certifi.where()
            )
            
            print(f"[ACTIVATION] Response status code: {response.status_code}")
            print(f"[ACTIVATION] Response headers: {dict(response.headers)}")
            print(f"[ACTIVATION] Raw response text: {response.text}")
            print(f"[ACTIVATION] Response status: {response.status_code}")
            print(f"[ACTIVATION] Response content: {response.text}")
        except requests.exceptions.SSLError as ssl_err:
            print(f"[ACTIVATION] SSL Error: {ssl_err}")
            return False, f"SSL 证书验证失败: {str(ssl_err)}"
        except requests.exceptions.ConnectionError as conn_err:
            print(f"[ACTIVATION] Connection Error: {conn_err}")
            return False, f"连接服务器失败: {str(conn_err)}"
        except requests.exceptions.Timeout as timeout_err:
            print(f"[ACTIVATION] Timeout Error: {timeout_err}")
            return False, "连接服务器超时，请检查网络后重试。"
        
        # 解析服务器响应
        try:
            response_data = response.json()
            print(f"[ACTIVATION] Full server response: {response_data}")
        except json.JSONDecodeError as e:
            print(f"[ACTIVATION] Failed to parse server response: {e}")
            print(f"[ACTIVATION] Raw response: {response.text}")
            return False, "激活失败：服务器响应格式错误"
        
        # 检查HTTP状态码和错误信息
        if response.status_code != 200 or not response_data.get("success", False):
            error_code = response_data.get("error", "UnknownError")
            error_details = response_data.get("details", "")
            print(f"[ACTIVATION] Server error: {error_code}, Details: {error_details}")
            
            # 特殊处理某些错误
            if error_code == "InvalidSignature":
                return False, "激活失败：产品密钥签名无效。\nActivation failed: Invalid product key signature."
            elif error_code == "TransactionFailed":
                return False, f"激活失败：事务处理失败。详情：{error_details}"
            
            return False, get_bilingual_error_message(error_code)
            
        support_key = response_data.get("support_key")
        if not support_key:
            print("[ACTIVATION] No support key in successful response")
            return False, get_bilingual_error_message("InvalidRequest")
            
        # 写入support_key到激活文件
        print("[ACTIVATION] Writing support key to activation file")
        with open(ACTIVATION_FILE, 'wb') as f:
            f.write(support_key.encode('utf-8'))
            
        # 验证写入的激活文件
        print("[ACTIVATION] Verifying activation file")
        verification_result, status_data = verify_activation()
        status_code = status_data.get("status_code")
        
        # 如果验证失败，清理激活文件并返回错误
        if status_code != "Activated":
            if os.path.exists(ACTIVATION_FILE):
                os.remove(ACTIVATION_FILE)
                print("[ACTIVATION] Removed invalid activation file")
            error_message = "激活验证失败"
            if verification_result and isinstance(verification_result, dict):
                error_message = verification_result.get("error_message", error_message)
            return False, error_message
            
        print("[ACTIVATION] Activation successful")
        return True, "产品激活成功！请重启软件以使更改生效。"
        
    except requests.exceptions.SSLError as ssl_err:
        print(f"[ACTIVATION] SSL Error: {str(ssl_err)}")
        print("[ACTIVATION] SSL Certificate verification failed")
        return False, f"SSL证书验证失败，请检查网络连接。\nSSL verification failed: {str(ssl_err)}"
    except requests.exceptions.Timeout:
        print("[ACTIVATION] Request timeout")
        print("[ACTIVATION] Connection timed out after 30 seconds")
        return False, get_bilingual_error_message("TimeoutError")
    except requests.exceptions.ConnectionError as conn_err:
        print(f"[ACTIVATION] Connection Error: {str(conn_err)}")
        print("[ACTIVATION] Failed to establish connection to the server")
        return False, "无法连接到激活服务器，请检查网络连接。\nCannot connect to activation server."
    except requests.exceptions.RequestException as e:
        print(f"[ACTIVATION] Network error: {str(e)}")
        print(f"[ACTIVATION] Request failed with error: {type(e).__name__}")
        return False, get_bilingual_error_message("NetworkError")
    except Exception as e:
        print(f"[ACTIVATION] Unexpected error: {str(e)}")
        print(f"[ACTIVATION] Error type: {type(e).__name__}")
        print(f"[ACTIVATION] Error details: {traceback.format_exc()}")
        return False, get_bilingual_error_message("UnknownError")

def activate_with_support_key(support_key_b64: str):
    """
    使用预加密的支持密钥进行激活。
    验证密钥的有效性并确保其与当前机器匹配。
    
    Args:
        support_key_b64 (str): 预加密的支持密钥（Base64格式）
        
    Returns:
        tuple: (success: bool, message: str)
    """
    try:
        print("[SUPPORT_KEY] 开始验证支持密钥")
        print(f"[SUPPORT_KEY] 支持密钥长度: {len(support_key_b64)}")
        
        # 确保输入是字符串格式
        if isinstance(support_key_b64, bytes):
            print("[SUPPORT_KEY] 输入是字节类型，转换为字符串")
            support_key_b64 = support_key_b64.decode('utf-8')
            
        # 处理可能的Base64填充问题
        # 先删除所有可能的填充符
        normalized_key = support_key_b64.rstrip('=')
        # 添加正确的填充
        normalized_key = normalized_key + '=' * (-len(normalized_key) % 4)
        print(f"[SUPPORT_KEY] 规范化后密钥长度: {len(normalized_key)}")
            
        # 使用 Fernet 解密并验证密钥
        print("[SUPPORT_KEY] 尝试解密密钥")
        fernet = Fernet(SECRET_KEY)
        try:
            # 先尝试直接解密原始输入
            try:
                print("[SUPPORT_KEY] 尝试方法1: 直接解密原始输入")
                decrypted_data = fernet.decrypt(support_key_b64.encode('utf-8'))
                print("[SUPPORT_KEY] 方法1成功")
            except Exception as e1:
                print(f"[SUPPORT_KEY] 方法1失败: {str(e1)}")
                # 尝试使用规范化的密钥
                try:
                    print("[SUPPORT_KEY] 尝试方法2: 使用规范化密钥")
                    decrypted_data = fernet.decrypt(normalized_key.encode('utf-8'))
                    print("[SUPPORT_KEY] 方法2成功")
                except Exception as e2:
                    print(f"[SUPPORT_KEY] 方法2失败: {str(e2)}")
                    # 尝试先解码Base64再解密
                    try:
                        print("[SUPPORT_KEY] 尝试方法3: 解码Base64后解密")
                        raw_data = base64.urlsafe_b64decode(normalized_key.encode('utf-8'))
                        # 如果解码成功，直接写入文件并验证
                        print("[SUPPORT_KEY] Base64解码成功，直接写入文件")
                        with open(ACTIVATION_FILE, 'wb') as f:
                            f.write(raw_data)
                        # 尝试验证
                        verification_result, status_data = verify_activation()
                        status_code = status_data.get("status_code")
                        if status_code == "Activated":
                            print("[SUPPORT_KEY] 直接写入解码数据激活成功")
                            return True, "产品已通过支持密钥成功激活！请重启软件。"
                        else:
                            # 清理无效文件
                            if os.path.exists(ACTIVATION_FILE):
                                os.remove(ACTIVATION_FILE)
                            print(f"[SUPPORT_KEY] 直接写入解码数据激活失败: {status_code}")
                            # 继续尝试其他方法
                            raise Exception("验证失败，尝试其他方法")
                    except Exception as e3:
                        print(f"[SUPPORT_KEY] 方法3失败: {str(e3)}")
                        # 尝试处理可能的 Node.js fernet 库生成的密钥
                        try:
                            print("[SUPPORT_KEY] 尝试方法4: 使用兼容模式处理可能的Node.js密钥")
                            # 这是另一种兼容性方法
                            with open(ACTIVATION_FILE, 'wb') as f:
                                f.write(support_key_b64.encode('utf-8'))
                            # 尝试验证
                            verification_result, status_data = verify_activation()
                            status_code = status_data.get("status_code")
                            if status_code == "Activated":
                                print("[SUPPORT_KEY] 方法4激活成功")
                                return True, "产品已通过支持密钥成功激活！请重启软件。"
                            else:
                                # 清理无效文件
                                if os.path.exists(ACTIVATION_FILE):
                                    os.remove(ACTIVATION_FILE)
                                print("[SUPPORT_KEY] 所有方法均失败")
                                return False, "激活失败：支持密钥无效或已损坏。"
                        except Exception as e4:
                            print(f"[SUPPORT_KEY] 方法4失败: {str(e4)}")
                            return False, "激活失败：支持密钥无效或已损坏。所有解密尝试均失败。"
                
            # 如果解密成功，继续处理
            activation_info = json.loads(decrypted_data.decode('utf-8'))
            print(f"[SUPPORT_KEY] 成功解密密钥并解析JSON数据")
            
            # 验证机器ID
            current_machine_id = get_machine_id()
            print(f"[SUPPORT_KEY] 当前机器ID: {current_machine_id}")
            print(f"[SUPPORT_KEY] 密钥机器ID: {activation_info.get('machine_id')}")
            if activation_info.get("machine_id") != current_machine_id:
                print("[SUPPORT_KEY] 机器ID不匹配")
                return False, "激活失败：此支持密钥已绑定至其他设备。"
                
            # 写入原始加密密钥到激活文件
            print("[SUPPORT_KEY] 写入激活文件")
            with open(ACTIVATION_FILE, 'wb') as f:
                # 使用解密成功的方法对应的密钥
                if 'raw_data' in locals():
                    print("[SUPPORT_KEY] 写入解码后的原始数据")
                    f.write(raw_data)
                else:
                    print("[SUPPORT_KEY] 写入原始支持密钥")
                    f.write(support_key_b64.encode('utf-8'))
                
            # 再次验证激活文件
            verification_result, status_data = verify_activation()
            status_code = status_data.get("status_code")
            if status_code == "Activated":
                print("[SUPPORT_KEY] 激活成功")
                return True, "产品已通过支持密钥成功激活！请重启软件。"
            else:
                # 清理无效文件
                if os.path.exists(ACTIVATION_FILE):
                    os.remove(ACTIVATION_FILE)
                print(f"[SUPPORT_KEY] 验证失败: {status_code}")
                return False, "激活失败：支持密钥验证未通过。"
        except Exception as e:
            print(f"[SUPPORT_KEY] 密钥解密或验证失败: {str(e)}")
            return False, "激活失败：支持密钥无效或已损坏。"
            
    except Exception as e:
        print(f"[SUPPORT_KEY] 意外错误: {str(e)}")
        print("[SUPPORT_KEY] 详细错误信息: ", traceback.format_exc())
        return False, f"激活失败：{str(e)}"

def check_trial_status():
    """
    Check trial status
    Returns: (is_valid: bool, message: str)
    """
    print("[TRIAL_CHECK] Starting trial status check...")
    try:
        # Check local trial file
        if not os.path.exists(TRIAL_FILE):
            print("[TRIAL_CHECK] No local trial file found. Requesting trial status from server...")
            # Call startTrial API
            machine_id = get_machine_id()
            print(f"[TRIAL_CHECK] Using machine ID: {machine_id}")
            
            # 发送试用请求到服务器
            payload = {
                "machineId": machine_id
            }
            print(f"[TRIAL_CHECK] Sending payload: {payload}")
            
            response = requests.post(
                FIREBASE_TRIAL_URL,
                json=payload,
                headers=COMMON_HEADERS,
                timeout=30,
                verify=certifi.where()
            )
            
            print(f"[TRIAL_CHECK] Server response status: {response.status_code}")
            print(f"[TRIAL_CHECK] Server response: {response.text}")
            
            if response.status_code != 200:
                error_data = response.json()
                error_code = error_data.get("error", "UnknownError")
                error_message = get_bilingual_error_message(error_code)
                print(f"[TRIAL_CHECK] Server error: {error_code} - {error_message}")
                return False, error_message
                
            data = response.json()
            expiry_date = data.get("expiryDate")
            
            if not expiry_date:
                print("[TRIAL_CHECK] No expiry date in server response")
                return False, get_bilingual_error_message("InvalidRequest")
                
            print(f"[TRIAL_CHECK] Received expiry date: {expiry_date}")
            # Save trial information
            with open(TRIAL_FILE, 'w') as f:
                f.write(expiry_date)
            print("[TRIAL_CHECK] Saved trial information to local file")
        else:
            print("[TRIAL_CHECK] Found existing trial file")
            # Read existing trial information
            with open(TRIAL_FILE, 'r') as f:
                expiry_date = f.read().strip()
            print(f"[TRIAL_CHECK] Read expiry date from file: {expiry_date}")
        
        # 修复时区比较问题
        current_time = datetime.utcnow().replace(tzinfo=timezone.utc)  # 确保当前时间是UTC时区
        expiry_time = datetime.fromisoformat(expiry_date.replace('Z', '+00:00'))
        print(f"[TRIAL_CHECK] Current time (UTC): {current_time}, Expiry time: {expiry_time}")
        
        if current_time >= expiry_time:
            print("[TRIAL_CHECK] Trial has expired")
            return False, get_bilingual_error_message("TrialExpired")
        else:
            remaining = expiry_time - current_time
            # 转换到用户本地时区
            local_expiry_time = expiry_time.astimezone(None)
            # 计算剩余天数和小时数
            remaining_days = remaining.days
            remaining_hours = remaining.seconds // 3600
            # 创建用户友好的消息
            friendly_message = f"试用剩余: {remaining_days} 天 {remaining_hours} 小时 (至 {local_expiry_time.strftime('%Y-%m-%d %H:%M:%S')})"
            print(f"[TRIAL_CHECK] Trial is valid. {remaining_days} days {remaining_hours} hours remaining")
            return True, friendly_message
            
    except requests.exceptions.RequestException as e:
        print(f"[TRIAL_CHECK] Network error: {str(e)}")
        return False, get_bilingual_error_message("NetworkError")
    except Exception as e:
        print(f"[TRIAL_CHECK] Unexpected error: {str(e)}")
        return False, get_bilingual_error_message("UnknownError")

def get_license_status():
    """Get current license status, including trial status check"""
    activation_info, status_data = verify_activation()
    status_code = status_data.get("status_code")

    if status_code == "Activated" and activation_info:
        license_details = activation_info.get("license_details", {})
        license_type = license_details.get("type", "Standard")
        return "Activated", {"license_type": license_type}
    
    # Check trial status
    is_trial_valid, trial_message = check_trial_status()
    if is_trial_valid:
        return "Trial", {"message": trial_message}
    
    # If trial is also invalid, return needs activation status
    error_message = "License activation required or invalid."
    if activation_info and isinstance(activation_info, dict):
        error_message = activation_info.get("error_message", error_message)
    elif status_data and "error_message" in status_data:
        error_message = status_data.get("error_message", error_message)

    return "NeedsActivation", {
        "status_code": status_code,
        "error_message": error_message,
        "trial_message": trial_message
    }

# 添加离线激活支持
OFFLINE_MODE = True  # 设置为 True 以启用离线激活