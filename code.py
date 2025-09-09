# code.py (V4.3.4 - Preview spacing synchronization fix)
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import time
import pandas as pd
import logging
import re
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
import threading
from typing import List, Tuple, Optional, Dict, Any
from tkinter import colorchooser
import json
from premailer import transform
import webbrowser
import traceback # 用于调试
import html
import random  # 添加random模块用于生成随机延迟时间

try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
    print("[IMPORT_CHECK] cryptography.fernet imported successfully.")
except ImportError:
    CRYPTO_AVAILABLE = False
    print("[IMPORT_CHECK] WARNING: cryptography.fernet could not be imported. Password encryption/decryption will be unavailable.")


# --- 授权系统集成 (更新导入) ---
from license_manager import (
    get_license_status, get_machine_id,
    activate_product, activate_with_support_key,
    PRODUCT_KEY_PREFIX,
)

# Configure logging (不变)
logging.basicConfig(
    filename="email_sending.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Language dictionaries (MODIFIED for delay labels)
LANGUAGES = {
    "zh": {
        "title": "优邮 Pro", "input_label": "输入设置:", "status_label": "运行状态:",
        "activation_dialog_title": "软件激活 - 优邮 Pro", "activation_needed_default_message": "请输入您的激活码以使用软件。",
        "machine_id_label": "本机机器码:", "enter_key_label": "激活码:", "activate_button_text": "激活",
        "activation_key_empty_error": "激活码不能为空。", "activation_success_message": "激活成功！软件将继续运行。",
        "activation_failed_message": "激活失败。", "license_activated_type": "产品已激活 ({license_type})",
        "license_expired": "您的授权已过期，请购买新授权或联系支持。", "license_mismatch": "此激活码已绑定至其他设备。如需换机，请联系客服。",
        "license_corrupted": "授权文件已损坏或无效，请尝试重新激活或联系支持。", "needs_activation_no_key": "软件尚未激活，请输入激活码。",
        "sender_label": "发件人邮箱:", "password_label": "邮箱密码:", "smtp_host_label": "SMTP 服务器:", "smtp_port_label": "SMTP 端口:",
        "subject_label": "邮件主题:", "cc_email_label": "抄送邮箱:", 
        # "delay_label": "发送间隔(秒):",  # 删除单个间隔
        "min_delay_label": "最小间隔(秒):", # 添加最小间隔
        "max_delay_label": "最大间隔(秒):", # 添加最大间隔
        "invalid_delay_range": "发送间隔设置错误：最小间隔不能大于最大间隔。", # 添加间隔范围错误提示
        "batch_size_label": "每批数量:",
        "batch_interval_label": "每组间隔(分钟):", "excel_label": "选择 Excel 文件:", "body_label": "邮件正文 (HTML - 可用 {占位符}):",
        "select_excel": "浏览", "language_label": "语言:", "start_button": "开始发送", "bold": "加粗", "italic": "斜体",
        "font_size": "字体大小:", "font_label": "字体:", "color_label": "文字颜色:", "bg_color_label": "背景颜色:",
        "attach_label": "上传附件:", "select_attach": "添加附件", "success_count_label": "发送成功数量:",
        "no_excel": "错误：项目文件夹中没有 .xlsx 文件", "invalid_number": "请输入有效的数字", "select_excel_prompt": "请选择一个 Excel 文件",
        "sending_to": "正在发送邮件到 {recipient}，抄送 {cc}...", "success": "成功发送邮件到 {recipient}，抄送 {cc}",
        "failed": "发送邮件到 {recipient}（抄送 {cc}）失败: {error}", "invalid_cc": "抄送邮箱格式错误: {cc_email}",
        "smtp_connecting": "尝试连接到 {smtp_host}:{smtp_port}", "smtp_login": "尝试登录用户: {sender}",
        "smtp_success": "成功连接到 SMTP 服务器", "smtp_failed": "SMTP 连接或登录失败: {error}", "smtp_closed": "已断开 SMTP 服务器连接",
        "invalid_email": "第 {row} 行邮箱无效: {email}", "invalid_email_format": "第 {row} 行邮箱格式错误: {email}",
        "row_failed": "处理第 {row} 行失败: {error}", "no_contacts": "无有效联系人，退出",
        "load_excel_failed": "加载 Excel 文件 {excel_file} 失败: {error}", "completed": "发送完成，共尝试向 {count} 个联系人发送邮件，抄送 {cc_email}",
        "used_excel": "使用的 Excel 文件: {excel_file}", "details_log": "详情见 email_sending.log",
        "batch_info": "正在发送第 {batch} 批（共 {total} 批），本批 {count} 个邮箱", "batch_wait": "本批发送完成，等待 {minutes} 分钟后发送下一批...",
        "save_settings": "保存设置", "save_settings_success": "保存设置成功", "help_button": "帮助", "clear_attachments": "清除附件", "attachments_selected": "已选择附件: {filenames}",
        "apply_font_button": "应用字体", "signature_button": "个性签名", "clear_signature_button": "清除签名",
        "signature_selected_label": "签名: {filename}", "no_signature_selected": "签名: 无", "pause_button": "暂停", "resume_button": "继续",
        "start_from_current": "从当前进度继续", # 添加从当前进度继续选项
        "pausing_sending": "暂停发送...当前邮件发送完毕后或批次间歇时生效。", "resuming_sending": "恢复发送...",
        "paused_log_msg": "发送已暂停。SMTP 服务器已断开。", "paused_during_batch_wait": "在批次等待期间暂停发送。",
        "resumed_log_msg": "发送已恢复。", "smtp_reconnecting": "正在重新连接 SMTP 服务器...", "smtp_reconnected": "SMTP 服务器已重新连接。",
        "smtp_reconnect_failed": "SMTP 服务器重新连接失败: {error}", "send_fail_pause_title": "发送失败提醒",
        "send_fail_pause_message": "发送邮件失败。\n错误详情: {error_details}\n\n可能由于邮箱服务器限制，请检查网络或邮箱设置，并稍后再继续发送。",
        "personalization_engine_label": "个性化引擎", "scan_placeholders_button": "扫描占位符 (从主题/正文)",
        "placeholder_label": "邮件内占位符", "map_to_excel_column_label": "映射到Excel列", "fallback_option_label": "备选项",
        "no_excel_columns_loaded": "没有加载Excel表格，或者没有列数据", "no_placeholders_found": "未在主题或正文中找到占位符 (格式: {占位符名称})",
        "preview_engine_label": "个性化预览", "select_excel_row_for_preview": "选择Excel行预览:", "refresh_preview_button": "刷新预览",
        "preview_subject_label": "预览主题:",
        "preview_body_label": "预览正文：(占位符替换后，格式与输入一致)", # MODIFIED
        "mapped_to": "已映射到:", "unmapped": "未映射", "excel_loaded_columns_detected": "Excel已加载, 检测到列: {columns}",
        "suggested_mapping_notice": "提示: 软件已尝试为常见占位符 (如 {Name}) 预设映射。",
        "placeholder_scan_instruction": "请先在邮件主题和正文中输入您想使用的占位符 (例如 {Company}或{Product})，然后点击扫描按钮。",
        "permanent_activate_button": "永久激活",
        "permanently_activated_button": "已永久激活",
        "trial_remaining": "试用剩余: {days} 天 {hours} 小时 (至 {expiry_time})"
    },
    "en": {
        "title": "MailRoute Pro", "input_label": "Input Settings:", "status_label": "Running Status:",
        "activation_dialog_title": "Software Activation - MailRoute Pro", "activation_needed_default_message": "Please enter your activation key to use the software.",
        "machine_id_label": "Machine ID:", "enter_key_label": "Activation Key:", "activate_button_text": "Activate",
        "activation_key_empty_error": "Activation key cannot be empty.", "activation_success_message": "Activation successful! The software will now proceed.",
        "activation_failed_message": "Activation failed.", "license_activated_type": "Product Activated ({license_type})",
        "license_expired": "Your license has expired. Please purchase a new one or contact support.",
        "license_mismatch": "This activation key is bound to another device. Contact support for transfer.",
        "license_corrupted": "License file is corrupted or invalid. Try reactivating or contact support.",
        "needs_activation_no_key": "Software is not activated. Please enter your activation key.",
        "sender_label": "Sender Email:", "password_label": "Email Password:", "smtp_host_label": "SMTP Server:", "smtp_port_label": "SMTP Port:",
        "subject_label": "Email Subject:", "cc_email_label": "CC Email:", 
        # "delay_label": "Send Delay (seconds):", # 删除单个间隔
        "min_delay_label": "Min Delay (s):", # 添加最小间隔
        "max_delay_label": "Max Delay (s):", # 添加最大间隔
        "invalid_delay_range": "Send delay setting error: Min delay cannot be greater than max delay.", # 添加间隔范围错误提示
        "batch_size_label": "Batch Size:",
        "batch_interval_label": "Batch Interval (minutes):", "excel_label": "Select Excel File:", "body_label": "Email Body (HTML - use {Placeholders}):",
        "select_excel": "Browse", "language_label": "Language:", "start_button": "Start Sending", "bold": "Bold", "italic": "Italic",
        "font_size": "Font Size:", "font_label": "Font:", "color_label": "Text Color:", "bg_color_label": "Background Color:",
        "attach_label": "Upload Attachments:", "select_attach": "Add Attachments", "success_count_label": "Successful Sends:",
        "no_excel": "Error: No .xlsx files found in the project folder", "invalid_number": "Please enter a valid number",
        "select_excel_prompt": "Please select an Excel file", "sending_to": "Sending email to {recipient}, CC to {cc}...",
        "success": "Successfully sent email to {recipient}, CC to {cc}", "failed": "Failed to send email to {recipient} (CC to {cc}): {error}",
        "invalid_cc": "Invalid CC email format: {cc_email}", "smtp_connecting": "Connecting to {smtp_host}:{smtp_port}",
        "smtp_login": "Logging in as: {sender}", "smtp_success": "Successfully connected to SMTP server",
        "smtp_failed": "SMTP connection or login failed: {error}", "smtp_closed": "Disconnected from SMTP server",
        "invalid_email": "Invalid email in row {row}: {email}", "invalid_email_format": "Invalid email format in row {row}: {email}",
        "row_failed": "Failed to process row {row}: {error}", "no_contacts": "No valid contacts found, exiting",
        "load_excel_failed": "Failed to load Excel file {excel_file}: {error}", "completed": "Sending completed, attempted to send to {count} contacts, CC to {cc_email}",
        "used_excel": "Used Excel file: {excel_file}", "details_log": "Details in email_sending.log",
        "batch_info": "Sending batch {batch} of {total}, {count} emails in this batch", "batch_wait": "Batch completed, waiting {minutes} minutes before the next batch...",
        "save_settings": "Save Settings", "save_settings_success": "Settings saved successfully", "help_button": "Help", "clear_attachments": "Clear Attachments", "attachments_selected": "Selected attachments: {filenames}",
        "apply_font_button": "Apply Font", "signature_button": "Signature", "clear_signature_button": "Clear Signature",
        "signature_selected_label": "Signature: {filename}", "no_signature_selected": "Signature: None", "pause_button": "Pause",
        "resume_button": "Resume", 
        "start_from_current": "Continue From Current", # 添加从当前进度继续选项
        "pausing_sending": "Pausing email sending... Will take effect after the current email or during batch interval.",
        "resuming_sending": "Resuming email sending...",
        "paused_log_msg": "Sending paused. SMTP server disconnected.",
        "paused_during_batch_wait": "Sending paused during batch interval.", "resumed_log_msg": "Sending resumed.",
        "smtp_reconnecting": "Reconnecting to SMTP server...", "smtp_reconnected": "SMTP server reconnected.",
        "smtp_reconnect_failed": "SMTP server reconnection failed: {error}", "send_fail_pause_title": "Send Failure Alert",
        "send_fail_pause_message": "Failed to send email.\nError details: {error_details}\n\nThis might be due to email server restrictions. Please check your network/email settings and try resuming later.",
        "personalization_engine_label": "Personalization Engine", "scan_placeholders_button": "Scan Placeholders (from Subject/Body)",
        "placeholder_label": "Placeholder in Email", "map_to_excel_column_label": "Map to Excel Column", "fallback_option_label": "Fallback Text",
        "no_excel_columns_loaded": "Excel not loaded or no columns found", "no_placeholders_found": "No placeholders found in Subject or Body (Format: {PlaceholderName})",
        "preview_engine_label": "Personalization Preview", "select_excel_row_for_preview": "Select Excel Row for Preview:",
        "refresh_preview_button": "Refresh Preview", "preview_subject_label": "Preview Subject:",
        "preview_body_label": "Preview Body: (After placeholder replacement, formatting matches input)", # MODIFIED
        "mapped_to": "Mapped to:", "unmapped": "Unmapped", "excel_loaded_columns_detected": "Excel loaded, detected columns: {columns}",
        "suggested_mapping_notice": "Tip: Software has attempted to preset mappings for common placeholders (like {Name}).",
        "placeholder_scan_instruction": "Please first enter the placeholders you want to use in the email subject and body (e.g., {Company} or {Product}), then click the scan button.",
        "permanent_activate_button": "Permanent Activation",
        "permanently_activated_button": "Permanently Activated",
        "trial_remaining": "Trial Remaining: {days} days {hours} hours (until {expiry_time})"
    },
    "en": {
        "title": "MailRoute Pro", "input_label": "Input Settings:", "status_label": "Running Status:",
        "activation_dialog_title": "Software Activation - MailRoute Pro", "activation_needed_default_message": "Please enter your activation key to use the software.",
        "machine_id_label": "Machine ID:", "enter_key_label": "Activation Key:", "activate_button_text": "Activate",
        "activation_key_empty_error": "Activation key cannot be empty.", "activation_success_message": "Activation successful! The software will now proceed.",
        "activation_failed_message": "Activation failed.", "license_activated_type": "Product Activated ({license_type})",
        "license_expired": "Your license has expired. Please purchase a new one or contact support.",
        "license_mismatch": "This activation key is bound to another device. Contact support for transfer.",
        "license_corrupted": "License file is corrupted or invalid. Try reactivating or contact support.",
        "needs_activation_no_key": "Software is not activated. Please enter your activation key.",
        "sender_label": "Sender Email:", "password_label": "Email Password:", "smtp_host_label": "SMTP Server:", "smtp_port_label": "SMTP Port:",
        "subject_label": "Email Subject:", "cc_email_label": "CC Email:", 
        # "delay_label": "Send Delay (seconds):", # 删除单个间隔
        "min_delay_label": "Min Delay (s):", # 添加最小间隔
        "max_delay_label": "Max Delay (s):", # 添加最大间隔
        "invalid_delay_range": "Send delay setting error: Min delay cannot be greater than max delay.", # 添加间隔范围错误提示
        "batch_size_label": "Batch Size:",
        "batch_interval_label": "Batch Interval (minutes):", "excel_label": "Select Excel File:", "body_label": "Email Body (HTML - use {Placeholders}):",
        "select_excel": "Browse", "language_label": "Language:", "start_button": "Start Sending", "bold": "Bold", "italic": "Italic",
        "font_size": "Font Size:", "font_label": "Font:", "color_label": "Text Color:", "bg_color_label": "Background Color:",
        "attach_label": "Upload Attachments:", "select_attach": "Add Attachments", "success_count_label": "Successful Sends:",
        "no_excel": "Error: No .xlsx files found in the project folder", "invalid_number": "Please enter a valid number",
        "select_excel_prompt": "Please select an Excel file", "sending_to": "Sending email to {recipient}, CC to {cc}...",
        "success": "Successfully sent email to {recipient}, CC to {cc}", "failed": "Failed to send email to {recipient} (CC to {cc}): {error}",
        "invalid_cc": "Invalid CC email format: {cc_email}", "smtp_connecting": "Connecting to {smtp_host}:{smtp_port}",
        "smtp_login": "Logging in as: {sender}", "smtp_success": "Successfully connected to SMTP server",
        "smtp_failed": "SMTP connection or login failed: {error}", "smtp_closed": "Disconnected from SMTP server",
        "invalid_email": "Invalid email in row {row}: {email}", "invalid_email_format": "Invalid email format in row {row}: {email}",
        "row_failed": "Failed to process row {row}: {error}", "no_contacts": "No valid contacts found, exiting",
        "load_excel_failed": "Failed to load Excel file {excel_file}: {error}", "completed": "Sending completed, attempted to send to {count} contacts, CC to {cc_email}",
        "used_excel": "Used Excel file: {excel_file}", "details_log": "Details in email_sending.log",
        "batch_info": "Sending batch {batch} of {total}, {count} emails in this batch", "batch_wait": "Batch completed, waiting {minutes} minutes before the next batch...",
        "save_settings": "Save Settings", "save_settings_success": "Settings saved successfully", "help_button": "Help", "clear_attachments": "Clear Attachments", "attachments_selected": "Selected attachments: {filenames}",
        "apply_font_button": "Apply Font", "signature_button": "Signature", "clear_signature_button": "Clear Signature",
        "signature_selected_label": "Signature: {filename}", "no_signature_selected": "Signature: None", "pause_button": "Pause",
        "resume_button": "Resume", 
        "start_from_current": "Continue From Current", # 添加从当前进度继续选项
        "pausing_sending": "Pausing email sending... Will take effect after the current email or during batch interval.",
        "resuming_sending": "Resuming email sending...",
        "paused_log_msg": "Sending paused. SMTP server disconnected.",
        "paused_during_batch_wait": "Sending paused during batch interval.", "resumed_log_msg": "Sending resumed.",
        "smtp_reconnecting": "Reconnecting to SMTP server...", "smtp_reconnected": "SMTP server reconnected.",
        "smtp_reconnect_failed": "SMTP server reconnection failed: {error}", "send_fail_pause_title": "Send Failure Alert",
        "send_fail_pause_message": "Failed to send email.\nError details: {error_details}\n\nThis might be due to email server restrictions. Please check your network/email settings and try resuming later.",
        "personalization_engine_label": "Personalization Engine", "scan_placeholders_button": "Scan Placeholders (from Subject/Body)",
        "placeholder_label": "Placeholder in Email", "map_to_excel_column_label": "Map to Excel Column", "fallback_option_label": "Fallback Text",
        "no_excel_columns_loaded": "Excel not loaded or no columns found", "no_placeholders_found": "No placeholders found in Subject or Body (Format: {PlaceholderName})",
        "preview_engine_label": "Personalization Preview", "select_excel_row_for_preview": "Select Excel Row for Preview:",
        "refresh_preview_button": "Refresh Preview", "preview_subject_label": "Preview Subject:",
        "preview_body_label": "Preview Body: (After placeholder replacement, formatting matches input)", # MODIFIED
        "mapped_to": "Mapped to:", "unmapped": "Unmapped", "excel_loaded_columns_detected": "Excel loaded, detected columns: {columns}",
        "suggested_mapping_notice": "Hint: Common placeholders (e.g., {Name}) may be pre-mapped if corresponding Excel columns are found.",
        "placeholder_scan_instruction": "First, type placeholders (e.g., {Company} or {Product}) into your Email Subject and Body, then click the Scan button.",
        "permanent_activate_button": "Permanent Activation",
        "permanently_activated_button": "Permanently Activated",
        "trial_remaining": "Trial Remaining: {days} days {hours} hours (until {expiry_time})"
    }
}


def is_valid_email(email: str) -> bool:
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))

class ScrollableFrame(ttk.Frame): # Unchanged
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self, highlightthickness=0); scrollbar_y = ttk.Scrollbar(self, orient="vertical", command=canvas.yview); scrollbar_x = ttk.Scrollbar(self, orient="horizontal", command=canvas.xview); self.scrollable_area = ttk.Frame(canvas)
        self.scrollable_area.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_area, anchor="nw"); canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        canvas.grid(row=0, column=0, sticky="nsew"); scrollbar_y.grid(row=0, column=1, sticky="ns"); scrollbar_x.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1); self.grid_columnconfigure(0, weight=1)

def send_single_email(*args, **kwargs): # Unchanged; collapsed for brevity
    app_instance = args[7] # Assuming app_instance is the 8th argument
    lang = args[8] # Assuming lang is the 9th argument
    recipient = args[2] # Assuming recipient is the 3rd argument
    cc = args[3] # Assuming cc is the 4th argument
    try:
        smtp_server, sender, _, _, subject_personalized, body_personalized_html, attachment_paths, _, _, signature_image_path, signature_image_cid = args[:11]
        msg = MIMEMultipart('related') if signature_image_path and os.path.exists(signature_image_path) and signature_image_cid in body_personalized_html else MIMEMultipart('alternative')
        msg["From"] = sender; msg["To"] = recipient; msg["Cc"] = cc; msg["Subject"] = subject_personalized
        msg.attach(MIMEText(body_personalized_html, "html", "utf-8"))
        if signature_image_path and signature_image_cid and os.path.exists(signature_image_path) and signature_image_cid in body_personalized_html :
            try:
                with open(signature_image_path, "rb") as f: img_data = f.read()
                img = MIMEImage(img_data); img.add_header('Content-ID', f'<{signature_image_cid}>'); img.add_header('Content-Disposition', 'inline', filename=os.path.basename(signature_image_path)); msg.attach(img)
            except Exception as e: logging.error(f"Failed to attach signature image {signature_image_path}: {e}"); app_instance.log_to_status(f"Warning: Failed to attach signature {os.path.basename(signature_image_path)}: {str(e)}")
        for attachment_path in attachment_paths:
            if attachment_path and os.path.exists(attachment_path):
                try:
                    with open(attachment_path, "rb") as f: part = MIMEBase("application", "octet-stream"); part.set_payload(f.read())
                    encoders.encode_base64(part); part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}"); msg.attach(part)
                except Exception as e: logging.error(f"Failed to attach file {attachment_path}: {e}"); app_instance.log_to_status(f"Warning: Failed to attach {attachment_path}: {str(e)}")
        recipients_list = [recipient] + ([cc_email.strip() for cc_email in cc.split(',') if cc_email.strip() and is_valid_email(cc_email.strip())] if cc else [])
        if not recipient: logging.warning(f"Recipient email is empty, skipping."); app_instance.log_to_status(f"Recipient email is empty, skipping."); return False, "Recipient email was empty."
        smtp_server.sendmail(sender, recipients_list, msg.as_string())
        logging.info(f"Successfully sent email to {recipient}, CC to {cc}"); app_instance.log_to_status(lang["success"].format(recipient=recipient, cc=cc)); return True, None
    except Exception as e: error_str = str(e); logging.error(f"Failed to send email to {recipient} (CC to {cc}): {error_str}"); app_instance.log_to_status(lang["failed"].format(recipient=recipient, cc=cc, error=error_str)); return False, error_str

def send_emails_to_contacts(*args, **kwargs):
    """发送邮件到联系人列表，支持从特定行开始发送"""
    # 解析参数
    if len(args) >= 19:  # 新版本的参数列表，包含start_row
        sender, password, contacts_df, subject_template, body_template_html_raw, placeholder_mappings, placeholder_fallbacks, attachment_paths, cc_email_str, smtp_host, smtp_port, min_delay, max_delay, batch_size, batch_interval, app_instance, lang, success_counter_widget, pause_event, start_row = args
    else:  # 向后兼容
        sender, password, contacts_df, subject_template, body_template_html_raw, placeholder_mappings, placeholder_fallbacks, attachment_paths, cc_email_str, smtp_host, smtp_port, min_delay, max_delay, batch_size, batch_interval, app_instance, lang, success_counter_widget, pause_event = args
        start_row = 0
    
    current_overall_success_count = app_instance.current_success_sends_count  # 从应用程序实例获取初始成功计数
    
    # 处理抄送邮箱
    if cc_email_str:
        invalid_ccs = [cc_ for cc_ in cc_email_str.split(',') if cc_.strip() and not is_valid_email(cc_.strip())]
        if invalid_ccs:
            for invalid_cc in invalid_ccs: 
                logging.error(f"Invalid CC email format: {invalid_cc}")
                app_instance.log_to_status(lang["invalid_cc"].format(cc_email=invalid_cc))
            cc_email_str = ",".join([cc_.strip() for cc_ in cc_email_str.split(',') if cc_.strip() and is_valid_email(cc_.strip())])
    
    smtp_server = None
    contact_df_iterator_idx = start_row  # 从指定行开始
    
    # 保存当前处理行到应用程序实例
    app_instance.current_excel_row = contact_df_iterator_idx
    
    while contact_df_iterator_idx < len(contacts_df):
        # 暂停功能处理
        pause_event.wait()
        if not pause_event.is_set():
            if smtp_server:
                try: smtp_server.quit()
                except Exception: pass
            smtp_server = None
            app_instance.log_to_status(lang["paused_log_msg"])
            pause_event.wait()
            app_instance.log_to_status(lang["resumed_log_msg"])
            

            
        # 尝试连接SMTP服务器
        if not smtp_server:
            app_instance.log_to_status(lang["smtp_reconnecting"] if contact_df_iterator_idx > start_row else lang["smtp_connecting"].format(smtp_host=smtp_host, smtp_port=smtp_port))
            try:
                smtp_server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=20)
                app_instance.log_to_status(lang["smtp_login"].format(sender=sender))
                smtp_server.login(sender, password)
                logging.info("Successfully reconnected/connected to SMTP server")
                app_instance.log_to_status(lang["smtp_reconnected"] if contact_df_iterator_idx > start_row else lang["smtp_success"])
            except Exception as e: 
                logging.error(f"SMTP reconnection/connection or login failed: {e}")
                app_instance.log_to_status(lang["smtp_reconnect_failed"].format(error=str(e)) if contact_df_iterator_idx > start_row else lang["smtp_failed"].format(error=str(e)))
                app_instance.root.after(0, app_instance.email_sending_finished)
                return current_overall_success_count
                
        # 记录批次信息
        if (contact_df_iterator_idx == start_row) or (contact_df_iterator_idx > start_row and contact_df_iterator_idx % batch_size == 0 and getattr(app_instance, '_last_logged_batch_start_idx', -1) != contact_df_iterator_idx):
            batch_num_display = ((contact_df_iterator_idx - start_row) // batch_size) + 1
            total_batches_display = ((len(contacts_df) - start_row) + batch_size - 1) // batch_size
            current_batch_size_actual = min(batch_size, len(contacts_df) - contact_df_iterator_idx)
            app_instance.log_to_status(lang["batch_info"].format(batch=batch_num_display, total=total_batches_display, count=current_batch_size_actual))
            app_instance._last_logged_batch_start_idx = contact_df_iterator_idx
            
        # 获取当前行数据
        row_data = contacts_df.iloc[contact_df_iterator_idx]
        recipient = row_data.get("Email")
        
        # 保存当前处理行到应用程序实例 - 这里始终是当前正在处理的行
        app_instance.current_excel_row = contact_df_iterator_idx
        
        # 检查邮箱有效性
        if pd.isna(recipient) or not isinstance(recipient, str) or not recipient.strip():
            logging.warning(f"Invalid email data in row {contact_df_iterator_idx+1}: {recipient}")
            app_instance.log_to_status(lang["invalid_email"].format(row=contact_df_iterator_idx+1, email=recipient))
            contact_df_iterator_idx += 1
            # 更新为下一个要处理的行索引
            app_instance.current_excel_row = contact_df_iterator_idx
            continue
            
        recipient = str(recipient).strip()
        if not is_valid_email(recipient):
            logging.warning(f"Invalid email format in row {contact_df_iterator_idx+1}: {recipient}")
            app_instance.log_to_status(lang["invalid_email_format"].format(row=contact_df_iterator_idx+1, email=recipient))
            contact_df_iterator_idx += 1
            # 更新为下一个要处理的行索引
            app_instance.current_excel_row = contact_df_iterator_idx
            continue
            
        # 处理邮件正文中的占位符替换
        current_subject = subject_template
        current_body_raw_html_for_email_build = body_template_html_raw  # This `body_template_html_raw` is already HTML
        for placeholder_key, excel_col_name in placeholder_mappings.items():
            placeholder_tag = f"{{{placeholder_key}}}"
            value_to_insert = ""
            if excel_col_name and excel_col_name in row_data:
                cell_value = row_data[excel_col_name]
                if pd.notna(cell_value): value_to_insert = str(cell_value)
                if not value_to_insert.strip(): value_to_insert = ""
            if not value_to_insert: value_to_insert = placeholder_fallbacks.get(placeholder_key, "")
            current_subject = current_subject.replace(placeholder_tag, value_to_insert)
            # IMPORTANT: For actual email sending, personalization happens on the HTML structure
            current_body_raw_html_for_email_build = current_body_raw_html_for_email_build.replace(placeholder_tag, html.escape(value_to_insert))  # Ensure values inserted into HTML are escaped

        # 添加签名
        signature_html_tag = f'<p><img src="cid:{app_instance.signature_image_cid}"></p>' if app_instance.signature_image_path and os.path.exists(app_instance.signature_image_path) else ""
        combined_body_html_with_sig_placeholder = current_body_raw_html_for_email_build + signature_html_tag
        full_html_for_email = f'<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><style type="text/css">body {{ font-family: {app_instance.font_var.get()}; font-size: {app_instance.font_size_var.get()}px; }}</style></head><body>{combined_body_html_with_sig_placeholder}</body></html>'
        final_inlined_html = transform(full_html_for_email)
        app_instance.log_to_status(lang["sending_to"].format(recipient=recipient, cc=cc_email_str))
        
        # 发送邮件
        send_successful, error_message = send_single_email(
            smtp_server, sender, recipient, cc_email_str, current_subject, 
            final_inlined_html, attachment_paths, app_instance, lang, 
            app_instance.signature_image_path, app_instance.signature_image_cid
        )
        
        if send_successful:
            current_overall_success_count += 1
            app_instance.current_success_sends_count = current_overall_success_count  # 更新应用程序实例中的成功计数
            app_instance.root.after(0, lambda count=current_overall_success_count: success_counter_widget.config(text=f"{lang['success_count_label']} {count}"))
            contact_df_iterator_idx += 1
            
            # 更新为下一个要处理的行索引 - 关键修改，保存下一个要处理的行索引
            app_instance.current_excel_row = contact_df_iterator_idx
            
            # 生成随机延迟时间
            if min_delay >= 0 and max_delay >= 0 and min_delay <= max_delay:
                random_delay = random.uniform(min_delay, max_delay)
                logging.info(f"Random delay: {random_delay:.2f} seconds")
                time.sleep(random_delay)
            elif min_delay > 0:  # 兼容性处理，如果只有最小延迟有效
                time.sleep(min_delay)
        else:
            app_instance.root.after(0, lambda err=error_message: app_instance.trigger_pause_and_popup(err))
            if smtp_server:
                try: smtp_server.quit()
                except Exception: pass
            smtp_server = None
            
        # 处理批次间隔
        if contact_df_iterator_idx > start_row and contact_df_iterator_idx < len(contacts_df) and (contact_df_iterator_idx - start_row) % batch_size == 0:
            app_instance.log_to_status(lang["batch_wait"].format(minutes=batch_interval))
            if smtp_server:
                try: smtp_server.quit()
                except Exception: pass
            smtp_server = None
            wait_start_time = time.time()
            
            # 等待批次间隔时间
            while time.time() - wait_start_time < batch_interval * 60:
                # 检查是否暂停
                if not pause_event.is_set():
                    if not getattr(app_instance, '_paused_during_batch_wait_logged', False):
                        app_instance.log_to_status(lang["paused_during_batch_wait"])
                        app_instance._paused_during_batch_wait_logged = True
                    pause_event.wait()
                    if getattr(app_instance, '_paused_during_batch_wait_logged', False):
                        app_instance.log_to_status(lang["resumed_log_msg"])
                        app_instance._paused_during_batch_wait_logged = False
                    break
                    

                    
                time.sleep(0.5)
            app_instance._paused_during_batch_wait_logged = False
            
    # 关闭SMTP连接
    if smtp_server:
        try: smtp_server.quit()
        except Exception: pass
        logging.info("Disconnected (end of all batches)")
        app_instance.log_to_status(lang["smtp_closed"] + " (end of all batches)")
        
    return current_overall_success_count


def load_contacts(excel_file: str, app_instance, lang: dict) -> Optional[pd.DataFrame]: # Unchanged
    print(f"[LOAD_CONTACTS] Attempting to load contacts from: {excel_file}")
    try:
        df = pd.read_excel(excel_file)
        if "Email" not in df.columns:
            messagebox.showerror("Error", "Excel file must contain 'Email' column.")
            app_instance.log_to_status("Error: Excel file must contain 'Email' column.")
            print("[LOAD_CONTACTS] Error: 'Email' column missing.")
            return None
        app_instance.excel_column_headers = df.columns.tolist()
        app_instance.log_to_status(lang.get("excel_loaded_columns_detected", "Excel loaded, cols: {columns}").format(columns=", ".join(app_instance.excel_column_headers)))
        print(f"[LOAD_CONTACTS] Columns detected: {app_instance.excel_column_headers}")
        if app_instance.root and app_instance.root.winfo_exists():
            app_instance.root.after(0, app_instance.scan_and_setup_placeholder_ui)
            app_instance.root.after(0, app_instance.update_preview_row_selector)
        app_instance.log_to_status(f"Loaded contacts from {excel_file}. Found columns: {', '.join(df.columns)}")
        print(f"[LOAD_CONTACTS] Successfully loaded {len(df)} contacts.")
        return df
    except Exception as e:
        print(f"[LOAD_CONTACTS] CRITICAL ERROR loading Excel file {excel_file}: {e}\n{traceback.format_exc()}")
        logging.error(f"Failed to load Excel file {excel_file}: {e}", exc_info=True)
        app_instance.log_to_status(lang["load_excel_failed"].format(excel_file=excel_file, error=str(e)))
        app_instance.excel_column_headers = []
        if app_instance.root and app_instance.root.winfo_exists():
            app_instance.root.after(0, app_instance.scan_and_setup_placeholder_ui)
            app_instance.root.after(0, app_instance.update_preview_row_selector)
        return None

class EmailSenderApp:
    def __init__(self, root_window):
        print("[APP_INIT] Initializing EmailSenderApp...")
        self.root = root_window
        self.lang_code = "en"
        self.language_var = tk.StringVar(value=self.lang_code)

        self._set_default_config_values(initializing=True)

        self.root.title(self.lang["title"])
        self.root.geometry("1250x850"); self.root.configure(bg="#f0f4f8")

        self.sender_var = tk.StringVar(); self.password_var = tk.StringVar(); self.smtp_host_var = tk.StringVar(); self.smtp_port_var = tk.StringVar()
        self.subject_var = tk.StringVar(); self.cc_email_var = tk.StringVar()
        # 将单一的delay_var替换为min_delay_var和max_delay_var
        self.min_delay_var = tk.StringVar(value="1")  
        self.max_delay_var = tk.StringVar(value="5")
        self.batch_size_var = tk.StringVar(value="100")
        self.batch_interval_var = tk.StringVar(value="10"); self.excel_file_var = tk.StringVar(); self.attachment_paths = []; self.attachment_var = tk.StringVar()
        self.font_size_var = tk.StringVar(value="14"); self.font_var = tk.StringVar(value="Arial")
        self.current_success_sends_count = 0; self.signature_image_path = None; self.signature_image_cid = "GlobalSignatureCID01"; self.signature_filename_var = tk.StringVar()
        self.fonts = ["Arial", "Helvetica", "Times New Roman", "Courier New", "Verdana", "宋体", "黑体", "楷体", "微软雅黑", "仿宋"]
        self.config_file = "config.json"; self.sending_thread = None; self.pause_event = threading.Event(); self.pause_event.set(); self.pause_resume_button = None
        
        # 添加帮助网址
        self.help_url = "https://help.getzentools.com/"
        
        # 添加当前Excel处理进度变量
        self.current_excel_row = 0  # 当前处理到的Excel行索引


        self.load_config()
        print(f"[APP_INIT] Config loaded. Language var from config: {self.language_var.get()}")

        self.lang = LANGUAGES.get(self.language_var.get(), LANGUAGES["en"])
        self.lang_code = self.language_var.get()
        print(f"[APP_INIT] Language finalized to: {self.lang_code}")
        self.root.title(self.lang["title"])

        print("[APP_INIT] Creating widgets...")
        self.create_widgets()
        print("[APP_INIT] Widgets created.")
        
        print("[APP_INIT] Applying loaded config to UI...")
        self.apply_loaded_config_to_ui()
        print("[APP_INIT] Loaded config applied to UI.")
        
        print("[APP_INIT] Checking license on startup...")
        self.check_license_on_startup() # This might withdraw and then show dialog
        # Ensure all pending UI updates are processed before mainloop starts
        if self.root.winfo_exists():
            self.root.update_idletasks()
            print("[APP_INIT] Forced root window update_idletasks after license check.")
        print("[APP_INIT] EmailSenderApp initialization complete.")

    def _set_default_config_values(self, initializing=False): # 更新默认配置值
        print("[CONFIG_HELPER] Setting default config values...")
        if not initializing:
            self.password_var.set("")
            self.sender_var.set("")
            self.smtp_host_var.set("")
            self.smtp_port_var.set("")
            self.subject_var.set("")
            self.cc_email_var.set("")
            # 使用min_delay_var和max_delay_var替代delay_var
            self.min_delay_var.set("1")
            self.max_delay_var.set("5")
            self.batch_size_var.set("100")
            self.batch_interval_var.set("10")
            self.excel_file_var.set("")
            self.attachment_paths = []
            self.attachment_var.set("")
            if not self.language_var.get():
                 self.language_var.set(self.lang_code)
            self.font_size_var.set("14")
            self.font_var.set("Arial")
            self.signature_image_path = None
        self.loaded_body_content = ""
        self.loaded_excel_data_preview = None
        self.excel_column_headers = []
        self.placeholder_map_frame_content = None
        self.placeholder_widgets = {}
        self.placeholder_config_mappings = {}
        self.placeholder_fallback_texts = {}
        current_lang_code = self.language_var.get() if self.language_var.get() else self.lang_code
        self.lang = LANGUAGES.get(current_lang_code, LANGUAGES["en"])
        if hasattr(self, 'signature_filename_var'):
             self.signature_filename_var.set(self.lang.get("no_signature_selected", "Signature: None"))

    def load_config(self): # 更新加载配置方法，处理旧的delay配置
        print("[CONFIG_LOAD] Starting to load config.")
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f: config = json.load(f)
                print("[CONFIG_LOAD] config.json read and parsed.")
                self.language_var.set(config.get("language", self.lang_code))
                self.lang_code = self.language_var.get()
                self.lang = LANGUAGES.get(self.lang_code, LANGUAGES["en"])
                print(f"[CONFIG_LOAD] Language from config set to: {self.lang_code}")
                pwd_enc_key_str = config.get("password_encryption_key"); encrypted_pwd_str = config.get("password_encrypted")
                if CRYPTO_AVAILABLE and pwd_enc_key_str and encrypted_pwd_str:
                    print("[CONFIG_LOAD] Found encrypted password and key in config. Attempting decryption.")
                    try:
                        key_bytes = pwd_enc_key_str.encode('utf-8'); config_fernet = Fernet(key_bytes)
                        decrypted_password = config_fernet.decrypt(encrypted_pwd_str.encode('utf-8')).decode('utf-8')
                        self.password_var.set(decrypted_password); print("[CONFIG_LOAD] Password decrypted successfully.")
                    except Exception as e_dcrypt: print(f"[CONFIG_LOAD] Failed to decrypt password from config: {e_dcrypt}"); logging.error(f"Failed to decrypt password: {e_dcrypt}", exc_info=True); self.password_var.set("")
                elif "password" in config: self.password_var.set(config.get("password", "")); print("[CONFIG_LOAD] Loaded unencrypted password from old config.")
                else: self.password_var.set(""); print("[CONFIG_LOAD] No password (encrypted or unencrypted) found in config.")
                
                # 设置其他配置项
                self.sender_var.set(config.get("sender", ""))
                self.smtp_host_var.set(config.get("smtp_host", ""))
                self.smtp_port_var.set(config.get("smtp_port", ""))
                self.subject_var.set(config.get("subject", ""))
                self.cc_email_var.set(config.get("cc_email", ""))
                
                # 处理延迟配置，支持新旧格式
                if "min_delay" in config and "max_delay" in config:
                    # 新格式配置已存在
                    self.min_delay_var.set(config.get("min_delay", "1"))
                    self.max_delay_var.set(config.get("max_delay", "5"))
                    print("[CONFIG_LOAD] Loaded min/max delay from config.")
                elif "delay" in config:
                    # 旧格式配置，单一延迟值，将其设置为最小和最大延迟
                    delay_value = config.get("delay", "1")
                    self.min_delay_var.set(delay_value)
                    self.max_delay_var.set(delay_value)
                    print(f"[CONFIG_LOAD] Migrated single delay value ({delay_value}) to min/max delay.")
                else:
                    # 无延迟配置，使用默认值
                    self.min_delay_var.set("1")
                    self.max_delay_var.set("5")
                    print("[CONFIG_LOAD] No delay config found, using defaults.")
                
                # 设置其他剩余配置项
                self.batch_size_var.set(config.get("batch_size", "100"))
                self.batch_interval_var.set(config.get("batch_interval", "10"))
                self.excel_file_var.set(config.get("excel_file", ""))
                self.attachment_paths = config.get("attachments", [])
                self.font_size_var.set(config.get("font_size", "14"))
                self.font_var.set(config.get("font", "Arial"))
                self.signature_image_path = config.get("signature_image_path")
                self.loaded_body_content = config.get("email_body", "")
                self.placeholder_config_mappings = config.get("placeholder_config_mappings", {})
                self.placeholder_fallback_texts = config.get("placeholder_fallback_texts", {})
                
                print("[CONFIG_LOAD] Other config values set from file.")
            except json.JSONDecodeError as e_json: print(f"[CONFIG_LOAD] Error decoding JSON from config file: {e_json}"); logging.error(f"Error decoding JSON from {self.config_file}: {e_json}"); self._set_default_config_values()
            except Exception as e: print(f"[CONFIG_LOAD] CRITICAL ERROR loading config file: {e}\n{traceback.format_exc()}"); logging.error(f"Failed to load config file {self.config_file}: {e}", exc_info=True); self._set_default_config_values()
        else: print("[CONFIG_LOAD] Config file not found. Using default values."); self._set_default_config_values()
        print("[CONFIG_LOAD] Config loading finished.")

    def check_license_on_startup(self):
        """
        检查许可证状态并相应地初始化UI。
        (升级版 v2 - 支持试用期和更详细的状态消息)
        """
        print("[LICENSE_CHECK] Starting license check.")
        if self.root.winfo_exists():
            self.root.withdraw()  # 先隐藏主窗口，直到确定状态

        try:
            # 关键！get_license_status 现在会先检查试用期
            status, data = get_license_status() 
            print(f"[LICENSE_CHECK] Status from license_manager: '{status}', Data: {data}")
        except Exception as e:
            print(f"[LICENSE_CHECK] CRITICAL ERROR calling get_license_status(): {e}\n{traceback.format_exc()}")
            logging.critical(f"CRITICAL ERROR getting license status: {e}", exc_info=True)
            if self.root.winfo_exists():
                messagebox.showerror("授权错误", f"无法检查授权状态: {e}\n应用程序将关闭。", parent=self.root if self.root.winfo_ismapped() else None)
                self.root.destroy()
            return

        # --- 根据状态决定下一步行动 ---

        if status == "Activated":
            # 激活状态：记录日志，显示主窗口
            license_type = data.get("license_type", self.lang.get("unknown", "未知类型"))
            display_message = self.lang.get("license_activated_type", "产品已激活 ({license_type})").format(license_type=license_type)
            self.log_to_status(display_message)
            print(f"[LICENSE_CHECK] Product is Activated ('{license_type}'). Deiconifying root window.")
            
            if self.root.winfo_exists():
                self.root.deiconify()
                
            # 更新激活按钮状态
            self.root.after(100, lambda: self.update_activate_button_state(status))

        elif status == "Trial":
            # 试用状态：记录试用信息，显示主窗口
            trial_message = data.get("message", "软件试用中。")
            self.log_to_status(trial_message)  # 在状态栏显示友好的试用信息
            print(f"[LICENSE_CHECK] In Trial period. Message: '{trial_message}'. Deiconifying root window.")
            if self.root.winfo_exists():
                self.root.deiconify()
                
            # 更新激活按钮状态
            self.root.after(100, lambda: self.update_activate_button_state(status))

        elif status == "NeedsActivation":
            # 需要激活状态：显示主窗口，然后在其上弹出激活对话框
            print(f"[LICENSE_CHECK] Product not activated (status: '{status}'). Showing activation dialog.")
            
            if self.root.winfo_exists():
                self.root.deiconify()
            
            self.root.after(0, lambda: self.show_activation_dialog(status, data))
            
            # 更新激活按钮状态
            self.root.after(100, lambda: self.update_activate_button_state(status))
            
        else:
            # 兜底情况
            messagebox.showerror("未知状态", f"收到未知的许可证状态: {status}\n应用程序将关闭。")
            if self.root.winfo_exists():
                self.root.destroy()

    def show_activation_dialog(self, lic_status_from_manager, data_from_manager): # Unchanged
        print(f"[ACTIVATION_DIALOG] Attempting to show activation dialog. Manager status: {lic_status_from_manager}")
        dialog = None
        try:
            # If root was withdrawn, dialog needs a parent that might not be visible yet.
            # Create dialog as Toplevel, it won't be visible until explicitly told or parent deiconified.
            dialog = tk.Toplevel(self.root) 
            dialog.title(self.lang.get("activation_dialog_title", "Activation Required"))
            dialog.geometry("480x280") # Let window manager place it initially
            dialog.resizable(False, False)
            dialog.transient(self.root) # Make it behave like a modal dialog for self.root
            
            # 确保激活对话框使用与主窗口相同的样式
            style = ttk.Style(dialog)
            
            # 定义颜色变量
            primary_color = "#2980b9"    # 主要按钮颜色
            success_color = "#27ae60"    # 成功按钮颜色（激活按钮）
            text_color_dark = "#2c3e50"  # 深色文本
            text_color_light = "#ffffff" # 浅色文本
            border_color = "#d0d9e3"     # 边框颜色
            
            # 确保Success.TButton样式在此对话框中可用
            style.configure("Success.TButton", 
                          font=("Helvetica", 10, "bold"),
                          background=success_color, 
                          foreground=text_color_dark,
                          borderwidth=1,
                          focusthickness=3,
                          padding=5)
            style.map("Success.TButton",
                    background=[("active", "#2ecc71"), ("disabled", "#cccccc")],
                    foreground=[("disabled", "#999999")])
            
            print("[ACTIVATION_DIALOG] Toplevel dialog window created.")
            dialog.update() # Force update to ensure the window is fully created and mapped
            print(f"[ACTIVATION_DIALOG] Dialog updated after creation. Mapped: {dialog.winfo_ismapped()}.")

            main_frame = ttk.Frame(dialog, padding="10"); main_frame.pack(fill="both", expand=True)
            status_code = data_from_manager.get("status_code", "")
            error_message_detail = data_from_manager.get("error_message", "")
            display_text_key = "activation_needed_default_message"
            if status_code == "NeedsActivation_NoKey": display_text_key = "needs_activation_no_key"
            elif status_code == "NeedsActivation_Mismatch": display_text_key = "license_mismatch"
            elif status_code == "NeedsActivation_Corrupted": display_text_key = "license_corrupted"
            elif status_code == "NeedsActivation_Expired": display_text_key = "license_expired"

            final_display_message = self.lang.get(display_text_key, "Activation Required (fallback).")
            if error_message_detail and display_text_key not in ["needs_activation_no_key", "activation_needed_default_message"]:
                 if final_display_message != error_message_detail: final_display_message = f"{final_display_message}\n({error_message_detail})"
            ttk.Label(main_frame, text=final_display_message, wraplength=440, justify="center", font=("Helvetica", 11)).pack(pady=(10, 20))

            machine_id = get_machine_id(); machine_id_frame = ttk.Frame(main_frame); machine_id_frame.pack(pady=10)
            ttk.Label(machine_id_frame, text=self.lang.get("machine_id_label", "Machine ID:")).pack(side="left")
            machine_id_entry = ttk.Entry(machine_id_frame, width=45); machine_id_entry.insert(0, machine_id); machine_id_entry.config(state="readonly"); machine_id_entry.pack(side="left", padx=5)
            ttk.Label(main_frame, text=self.lang.get("enter_key_label", "Activation Key:")).pack(pady=(10, 5))
            key_entry_var = tk.StringVar(); key_entry = ttk.Entry(main_frame, textvariable=key_entry_var, width=60, font=("Courier New", 9)); key_entry.pack(pady=5); key_entry.focus_set()
            print("[ACTIVATION_DIALOG] Dialog widgets created.")

            def attempt_activation_from_dialog():
                print("[ACTIVATION_DIALOG_ACTION] 'Activate' button clicked.")
                key_str = key_entry_var.get().strip()
                if not key_str:
                    messagebox.showerror(self.lang.get("activation_dialog_title"), self.lang.get("activation_key_empty_error"), parent=dialog)
                    return
                success, activation_msg = False, "Unknown activation error."
                try:
                    if key_str.startswith(PRODUCT_KEY_PREFIX + "-") and "." in key_str:
                        success, activation_msg = activate_product(key_str)
                    else:
                        success, activation_msg = activate_with_support_key(key_str)
                    print(f"[ACTIVATION_DIALOG_ACTION] Activation attempt result: Success={success}, Msg='{activation_msg}'")
                except Exception as e_act:
                    print(f"[ACTIVATION_DIALOG_ACTION] Error during activation call: {e_act}\n{traceback.format_exc()}")
                    success = False; activation_msg = f"Activation call error: {e_act}"

                if success:
                    # 更新激活按钮状态
                    if hasattr(self, 'activate_button') and self.activate_button.winfo_exists():
                        self.activate_button.config(text=self.lang["permanently_activated_button"], state=tk.DISABLED)
                    
                    messagebox.showinfo(self.lang.get("activation_dialog_title"), self.lang.get("activation_success_message"), parent=dialog)
                    if dialog and dialog.winfo_exists(): dialog.destroy() # Destroy dialog first
                    if self.root.winfo_exists(): self.root.deiconify() # Then show main window
                    print("[ACTIVATION_DIALOG_ACTION] Activation successful. Main window should be/become visible.")
                else:
                    messagebox.showerror(self.lang.get("activation_dialog_title"), f"{self.lang.get('activation_failed_message')}\n{activation_msg}", parent=dialog)
                    print("[ACTIVATION_DIALOG_ACTION] Activation failed.")
            
            def on_dialog_close_attempt():
                print("[ACTIVATION_DIALOG_ACTION] Dialog 'X' button clicked. Closing application.")
                if dialog and dialog.winfo_exists(): dialog.destroy()
                if self.root.winfo_exists(): self.root.destroy() # Ensure main app also closes

            dialog.protocol("WM_DELETE_WINDOW", on_dialog_close_attempt)
            # 为激活按钮应用Success.TButton样式
            ttk.Button(main_frame, text=self.lang.get("activate_button_text", "Activate"), 
                     command=attempt_activation_from_dialog, 
                     style="Success.TButton").pack(pady=20)
            
            # Center dialog relative to screen if root is not yet visible
            dialog.update_idletasks() # Allow tkinter to calculate sizes
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            size = tuple(int(_) for _ in dialog.geometry().split('+')[0].split('x'))
            x = screen_width/2 - size[0]/2
            y = screen_height/2 - size[1]/2
            dialog.geometry("+%d+%d" % (x, y))
            print(f"[ACTIVATION_DIALOG] Dialog positioned at +{int(x)}+{int(y)}. Size: {size[0]}x{size[1]}.")

            dialog.lift()
            print("[ACTIVATION_DIALOG] Dialog lifted.")
            # Removed grab_set and focus_force to debug hanging issue
            # dialog.grab_set() # Make it modal
            # print("[ACTIVATION_DIALOG] Dialog grab_set.")
            # dialog.focus_force() # Focus on dialog
            # print("[ACTIVATION_DIALOG] Dialog focus_force.")
            # dialog.wait_window() # Removed to debug hanging issue

            print("[ACTIVATION_DIALOG] Dialog setup complete.")

        except Exception as e:
            print(f"[ACTIVATION_DIALOG] CRITICAL ERROR creating/showing activation dialog: {e}")
            traceback.print_exc() # Print full traceback for debugging
            logging.critical(f"CRITICAL ERROR creating/showing activation dialog: {e}", exc_info=True)
            # Ensure root and dialog (if exists) are handled before exiting
            if dialog and dialog.winfo_exists():
                try: dialog.destroy()
                except tk.TclError: pass # Avoid error if already destroyed
            if self.root.winfo_exists():
                try:
                    messagebox.showerror("Activation Error", f"无法显示激活窗口: {e}\n应用程序将关闭。", parent=self.root if self.root.winfo_ismapped() else None)
                    self.root.destroy()
                except tk.TclError: pass # Avoid error if already destroyed
            return
    
    def log_to_status(self, message: str): # Unchanged
        if hasattr(self, 'status_text') and self.status_text and self.status_text.winfo_exists():
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S"); self.status_text.insert(tk.END, f"{timestamp} - {message}\n"); self.status_text.see(tk.END)
        else: print(f"[LOG_TO_STATUS_FALLBACK] Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')} - Message: {message}")

    def create_widgets(self): # Unchanged
        # 创建现代化、专业的UI主题
        style = ttk.Style()
        
        # 定义颜色变量
        primary_color = "#2980b9"    # 主要按钮颜色（开始按钮）
        success_color = "#27ae60"    # 成功按钮颜色（激活按钮）
        text_color_dark = "#2c3e50"  # 深色文本
        text_color_light = "#ffffff" # 浅色文本
        border_color = "#d0d9e3"     # 边框颜色
        disabled_bg = "#e9e9e9"      # 禁用状态背景
        hover_color = "#eaf2f8"      # 悬停颜色
        
        # 配置主要样式元素
        
        # 按钮样式
        style.configure("TButton", 
                      font=("Helvetica", 10),
                      background=primary_color, 
                      foreground=text_color_dark,
                      borderwidth=1,
                      focusthickness=3,
                      padding=5)
        style.map("TButton",
                background=[("active", "#3498db"), ("disabled", hover_color)],
                foreground=[("disabled", "#a0a0a0")])
        
        # 次要按钮样式
        style.configure("Secondary.TButton", 
                      font=("Helvetica", 10),
                      background="#ecf0f1", 
                      foreground=text_color_dark,
                      borderwidth=1,
                      focusthickness=3,
                      padding=5)
        style.map("Secondary.TButton",
                background=[("active", "#d6dce0"), ("disabled", disabled_bg)],
                foreground=[("disabled", "#a0a0a0")])
        
        # 成功按钮样式（用于激活按钮）
        style.configure("Success.TButton", 
                      font=("Helvetica", 10, "bold"),
                      background=success_color, 
                      foreground=text_color_dark,
                      borderwidth=1,
                      focusthickness=3,
                      padding=5)
        style.map("Success.TButton",
                background=[("active", "#2ecc71"), ("disabled", disabled_bg)],
                foreground=[("disabled", "#a0a0a0")])
        
        # 永久激活按钮样式
        style.configure("Activate.TButton", 
                      font=("Helvetica", 10, "bold"),
                      background="#f0f4f8", 
                      foreground="#006400",  # 深绿色
                      borderwidth=1,
                      focusthickness=3,
                      padding=5)
        style.map("Activate.TButton",
                background=[("active", "#e0e8e0"), ("disabled", disabled_bg)],
                foreground=[("disabled", "#a0a0a0")])
        
        # 标签样式
        style.configure("TLabel", 
                      font=("Helvetica", 10), 
                      foreground=text_color_dark)
        
        # 输入框样式
        style.configure("TEntry", 
                      font=("Helvetica", 10), 
                      fieldbackground=text_color_light,
                      foreground=text_color_dark,
                      bordercolor=border_color)
        style.map("TEntry", 
                fieldbackground=[("disabled", disabled_bg)],
                foreground=[("disabled", "#a0a0a0")])
        
        # 下拉框样式
        style.configure("TCombobox", 
                      font=("Helvetica", 10),
                      fieldbackground=text_color_light,
                      foreground=text_color_dark,
                      selectbackground=primary_color,
                      selectforeground=text_color_light)
        style.map("TCombobox",
                fieldbackground=[("readonly", text_color_light), ("disabled", disabled_bg)],
                selectbackground=[("readonly", primary_color)],
                selectforeground=[("readonly", text_color_light)])
        
        # 标签框样式
        style.configure("TLabelframe", 
                      borderwidth=1, 
                      relief=tk.SOLID,
                      bordercolor=border_color)
        style.configure("TLabelframe.Label", 
                      font=("Helvetica", 10, "bold"),
                      foreground=text_color_dark)
        
        # 创建其他UI元素
        main_paned_window = ttk.PanedWindow(self.root, orient=tk.VERTICAL); main_paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        top_frame_container = ttk.Frame(main_paned_window); main_paned_window.add(top_frame_container, weight=1); top_frame_container.columnconfigure(0, weight=1); top_frame_container.columnconfigure(1, weight=1); top_frame_container.rowconfigure(0, weight=1)
        self.input_frame = ttk.LabelFrame(top_frame_container, text=self.lang["input_label"], padding="10 20"); self.input_frame.grid(row=0, column=0, padx=(0,5), pady=5, sticky="nsew"); self.input_frame.columnconfigure(1, weight=1)
        ttk.Label(self.input_frame, text=self.lang["sender_label"]).grid(row=0, column=0, padx=5, pady=2, sticky="w"); ttk.Entry(self.input_frame, textvariable=self.sender_var, width=35).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Label(self.input_frame, text=self.lang["password_label"]).grid(row=1, column=0, padx=5, pady=2, sticky="w"); ttk.Entry(self.input_frame, textvariable=self.password_var, show="*", width=35).grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ttk.Label(self.input_frame, text=self.lang["smtp_host_label"]).grid(row=2, column=0, padx=5, pady=2, sticky="w"); ttk.Entry(self.input_frame, textvariable=self.smtp_host_var, width=35).grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        ttk.Label(self.input_frame, text=self.lang["smtp_port_label"]).grid(row=3, column=0, padx=5, pady=2, sticky="w"); ttk.Entry(self.input_frame, textvariable=self.smtp_port_var, width=35).grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        ttk.Label(self.input_frame, text=self.lang["subject_label"]).grid(row=4, column=0, padx=5, pady=2, sticky="w"); self.subject_entry = ttk.Entry(self.input_frame, textvariable=self.subject_var, width=35); self.subject_entry.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        ttk.Label(self.input_frame, text=self.lang["cc_email_label"]).grid(row=5, column=0, padx=5, pady=2, sticky="w"); ttk.Entry(self.input_frame, textvariable=self.cc_email_var, width=35).grid(row=5, column=1, padx=5, pady=2, sticky="ew")
        
        # 创建用于显示两个延迟输入框的Frame
        delay_frame = ttk.Frame(self.input_frame)
        delay_frame.grid(row=6, column=1, padx=5, pady=2, sticky="ew")
        delay_frame.columnconfigure(0, weight=1)  # 最小延迟输入框
        delay_frame.columnconfigure(2, weight=1)  # 最大延迟输入框
        
        # 添加最小延迟标签和输入框
        ttk.Label(self.input_frame, text=self.lang["min_delay_label"]).grid(row=6, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(delay_frame, textvariable=self.min_delay_var, width=15).grid(row=0, column=0, sticky="w")
        
        # 添加最大延迟标签和输入框
        ttk.Label(delay_frame, text=self.lang["max_delay_label"]).grid(row=0, column=1, padx=(20, 5), pady=2, sticky="w")
        ttk.Entry(delay_frame, textvariable=self.max_delay_var, width=15).grid(row=0, column=2, sticky="w")
        
        ttk.Label(self.input_frame, text=self.lang["batch_size_label"]).grid(row=7, column=0, padx=5, pady=2, sticky="w"); ttk.Entry(self.input_frame, textvariable=self.batch_size_var, width=35).grid(row=7, column=1, padx=5, pady=2, sticky="ew")
        ttk.Label(self.input_frame, text=self.lang["batch_interval_label"]).grid(row=8, column=0, padx=5, pady=2, sticky="w"); ttk.Entry(self.input_frame, textvariable=self.batch_interval_var, width=35).grid(row=8, column=1, padx=5, pady=2, sticky="ew")
        ttk.Label(self.input_frame, text=self.lang["excel_label"]).grid(row=9, column=0, padx=5, pady=2, sticky="w"); excel_frame = ttk.Frame(self.input_frame); excel_frame.grid(row=9, column=1, padx=5, pady=2, sticky="ew"); excel_frame.columnconfigure(0, weight=1); excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_file_var); excel_entry.grid(row=0, column=0, sticky="ew"); ttk.Button(excel_frame, text=self.lang["select_excel"], command=self.select_excel_file, style="Secondary.TButton").grid(row=0, column=1, padx=(5,0))
        ttk.Label(self.input_frame, text=self.lang["language_label"]).grid(row=10, column=0, padx=5, pady=2, sticky="w"); language_combo = ttk.Combobox(self.input_frame, textvariable=self.language_var, values=list(LANGUAGES.keys()), state="readonly", width=10); language_combo.grid(row=10, column=1, padx=5, pady=2, sticky="w"); language_combo.bind("<<ComboboxSelected>>", self.change_language)
        
        # 创建按钮框架，用于并排放置帮助按钮和保存设置按钮
        buttons_frame = ttk.Frame(self.input_frame)
        buttons_frame.grid(row=11, column=1, padx=5, pady=5, sticky="e")
        
        # 添加帮助按钮
        ttk.Button(buttons_frame, text=self.lang["help_button"], command=self.open_help_website, style="Secondary.TButton").pack(side="left", padx=(0,5))
        
        # 添加保存设置按钮
        ttk.Button(buttons_frame, text=self.lang["save_settings"], command=self.save_config, style="Secondary.TButton").pack(side="left")
        
        body_frame = ttk.LabelFrame(top_frame_container, text=self.lang["body_label"], padding="10"); body_frame.grid(row=0, column=1, padx=(5,0), pady=5, sticky="nsew"); body_frame.columnconfigure(0, weight=1); body_frame.rowconfigure(1, weight=1)
        toolbar_frame = ttk.Frame(body_frame); toolbar_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,5))
        ttk.Button(toolbar_frame, text=self.lang["bold"], command=self.toggle_bold, style="Secondary.TButton").pack(side="left", padx=2); ttk.Button(toolbar_frame, text=self.lang["italic"], command=self.toggle_italic, style="Secondary.TButton").pack(side="left", padx=2); ttk.Label(toolbar_frame, text=self.lang["font_size"]).pack(side="left", padx=(10,2)); ttk.Entry(toolbar_frame, textvariable=self.font_size_var, width=4).pack(side="left", padx=2); ttk.Label(toolbar_frame, text=self.lang["font_label"]).pack(side="left", padx=(10,2)); font_combo = ttk.Combobox(toolbar_frame, textvariable=self.font_var, values=self.fonts, state="readonly", width=12); font_combo.pack(side="left", padx=2); font_combo.bind("<<ComboboxSelected>>", self.apply_font_and_size_change); ttk.Button(toolbar_frame, text=self.lang["apply_font_button"], command=self.apply_font_and_size_change, style="Secondary.TButton").pack(side="left", padx=2); ttk.Button(toolbar_frame, text=self.lang["color_label"], command=self.choose_color, style="Secondary.TButton").pack(side="left", padx=2); ttk.Button(toolbar_frame, text=self.lang["bg_color_label"], command=self.choose_bg_color, style="Secondary.TButton").pack(side="left", padx=2)
        self.body_text_area = tk.Text(body_frame, height=10, width=40, font=(self.font_var.get(), int(self.font_size_var.get() or 14)), undo=True, wrap=tk.WORD); self.body_text_area.grid(row=1, column=0, sticky="nsew", padx=(0,0), pady=(0,0)); self.body_text_area.bind("<<Paste>>", self.handle_paste); v_scrollbar = ttk.Scrollbar(body_frame, orient=tk.VERTICAL, command=self.body_text_area.yview); v_scrollbar.grid(row=1, column=1, sticky="ns"); self.body_text_area.config(yscrollcommand=v_scrollbar.set)
        middle_frame_container = ttk.Frame(main_paned_window); main_paned_window.add(middle_frame_container, weight=4); middle_frame_container.columnconfigure(0, weight=1); middle_frame_container.columnconfigure(1, weight=1); middle_frame_container.rowconfigure(0, weight=1)
        self.placeholder_mapping_frame = ttk.LabelFrame(middle_frame_container, text=self.lang["personalization_engine_label"], padding="10"); self.placeholder_mapping_frame.grid(row=0, column=0, padx=(0,5), pady=5, sticky="nsew"); self.placeholder_mapping_frame.columnconfigure(0, weight=1); self.placeholder_mapping_frame.rowconfigure(1, weight=1)
        scan_button = ttk.Button(self.placeholder_mapping_frame, text=self.lang["scan_placeholders_button"], command=self.scan_and_setup_placeholder_ui, style="Secondary.TButton"); scan_button.grid(row=0, column=0, pady=(0,5), sticky="ew"); self.placeholder_scan_instruction_label = ttk.Label(self.placeholder_mapping_frame, text=self.lang["placeholder_scan_instruction"], wraplength=300, justify=tk.LEFT); self.placeholder_scan_instruction_label.grid(row=2, column=0, pady=(5,0), sticky="ew"); self.placeholder_scrollable_frame = ScrollableFrame(self.placeholder_mapping_frame); self.placeholder_scrollable_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        self.preview_frame = ttk.LabelFrame(middle_frame_container, text=self.lang["preview_engine_label"], padding="10"); self.preview_frame.grid(row=0, column=1, padx=(5,0), pady=5, sticky="nsew"); self.preview_frame.columnconfigure(0, weight=1); self.preview_frame.rowconfigure(1, weight=0); self.preview_frame.rowconfigure(2, weight=0); self.preview_frame.rowconfigure(3, weight=0); self.preview_frame.rowconfigure(4, weight=3)
        preview_controls_frame = ttk.Frame(self.preview_frame); preview_controls_frame.grid(row=0, column=0, sticky="ew", pady=(0,5)); preview_controls_frame.columnconfigure(1, weight=1); ttk.Label(preview_controls_frame, text=self.lang["select_excel_row_for_preview"]).grid(row=0, column=0, padx=(0,5), sticky="w"); self.preview_row_var = tk.StringVar(); self.preview_row_combobox = ttk.Combobox(preview_controls_frame, textvariable=self.preview_row_var, state="readonly", width=10); self.preview_row_combobox.grid(row=0, column=1, sticky="ew"); self.preview_row_combobox.bind("<<ComboboxSelected>>", self.update_preview); ttk.Button(preview_controls_frame, text=self.lang["refresh_preview_button"], command=self.update_preview, style="Secondary.TButton").grid(row=0, column=2, padx=(5,0), sticky="e")
        ttk.Label(self.preview_frame, text=self.lang["preview_subject_label"]).grid(row=1, column=0, sticky="w", pady=(5,2)); self.preview_subject_text = tk.Text(self.preview_frame, height=2, width=40, state="disabled", wrap=tk.WORD, bg="#e9e9e9"); self.preview_subject_text.grid(row=2, column=0, sticky="nsew", pady=(0,5));
        self.preview_body_label_widget = ttk.Label(self.preview_frame, text=self.lang["preview_body_label"]) # Store for easy update
        self.preview_body_label_widget.grid(row=3, column=0, sticky="w", pady=(5,2));
        self.preview_body_text = tk.Text(self.preview_frame, height=8, width=40, state="disabled", wrap=tk.WORD, bg="#e9e9e9"); self.preview_body_text.grid(row=4, column=0, sticky="nsew")
        bottom_frame_container = ttk.Frame(main_paned_window, height=200); main_paned_window.add(bottom_frame_container, weight=1); bottom_frame_container.columnconfigure(0, weight=1); bottom_frame_container.rowconfigure(0, weight=0); bottom_frame_container.rowconfigure(1, weight=1)
        button_controls_frame = ttk.Frame(bottom_frame_container); button_controls_frame.grid(row=0, column=0, pady=(10,5), sticky="ew")
        center_buttons_frame = ttk.Frame(button_controls_frame); center_buttons_frame.pack(side=tk.LEFT, padx=(20,0)); 
        self.start_button = ttk.Button(center_buttons_frame, text=self.lang["start_button"], command=self.start_sending); 
        self.start_button.pack(side="left", padx=(0, 5)); 
        self.pause_resume_button = ttk.Button(center_buttons_frame, text=self.lang["pause_button"], command=self.toggle_pause_resume, state=tk.DISABLED, style="Secondary.TButton"); 
        self.pause_resume_button.pack(side="left", padx=(0, 5))
        

        
        self.success_counter_label_widget = ttk.Label(center_buttons_frame, text=f"{self.lang['success_count_label']} 0")
        self.success_counter_label_widget.pack(side="left", padx=(5, 20))
        
        # 添加永久激活按钮
        self.activate_button = ttk.Button(center_buttons_frame, text=self.lang["permanent_activate_button"], command=self.open_activation_dialog, style="Activate.TButton")
        self.activate_button.pack(side="left", padx=(0, 20))
        
        right_controls_frame = ttk.Frame(button_controls_frame); right_controls_frame.pack(side=tk.RIGHT, padx=(0,20)); attach_label = ttk.Label(right_controls_frame, text=self.lang["attach_label"]); attach_label.pack(side="left", padx=(0,2)); self.attachment_entry = ttk.Entry(right_controls_frame, textvariable=self.attachment_var, width=25, state="readonly"); self.attachment_entry.pack(side="left", padx=(0,5)); ttk.Button(right_controls_frame, text=self.lang["select_attach"], command=self.select_attachments, style="Secondary.TButton").pack(side="left", padx=(0,2)); ttk.Button(right_controls_frame, text=self.lang["clear_attachments"], command=self.clear_attachments, style="Secondary.TButton").pack(side="left", padx=(0,10)); ttk.Button(right_controls_frame, text=self.lang["signature_button"], command=self.select_signature_image, style="Secondary.TButton").pack(side="left", padx=(5, 2)); ttk.Button(right_controls_frame, text=self.lang["clear_signature_button"], command=self.clear_signature_image, style="Secondary.TButton").pack(side="left", padx=(0,2)); self.signature_label_display = ttk.Label(right_controls_frame, textvariable=self.signature_filename_var, width=15, anchor="w"); self.signature_label_display.pack(side="left", padx=(0,5))
        status_frame = ttk.LabelFrame(bottom_frame_container, text=self.lang["status_label"], padding="10"); status_frame.grid(row=1, column=0, padx=0, pady=(5,0), sticky="nsew"); status_frame.columnconfigure(0, weight=1); status_frame.rowconfigure(0, weight=1); self.status_text = tk.Text(status_frame, height=6, width=70, state="normal", bg="#e8f4f8", wrap=tk.WORD); self.status_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.update_signature_display()

    def scan_and_setup_placeholder_ui(self): # Unchanged
        if hasattr(self, 'placeholder_map_frame_content') and self.placeholder_map_frame_content: self.placeholder_map_frame_content.destroy()
        self.placeholder_map_frame_content = ttk.Frame(self.placeholder_scrollable_frame.scrollable_area); self.placeholder_map_frame_content.pack(fill="x", expand=True)
        self.placeholder_map_frame_content.columnconfigure(0, weight=2); self.placeholder_map_frame_content.columnconfigure(1, weight=3); self.placeholder_map_frame_content.columnconfigure(2, weight=3)
        self.placeholder_widgets.clear(); subject_text = self.subject_var.get(); body_text = "";
        if hasattr(self, 'body_text_area') and self.body_text_area.winfo_exists(): body_text = self.body_text_area.get("1.0", tk.END)
        combined_text_to_scan = subject_text + " " + body_text; found_placeholders = sorted(list(set(re.findall(r"\{([a-zA-Z0-9_]+)\}", combined_text_to_scan))))
        if not self.excel_column_headers: ttk.Label(self.placeholder_map_frame_content, text=self.lang.get("no_excel_columns_loaded","Excel not loaded")).grid(row=0, column=0, columnspan=3, pady=10, sticky="ew"); self.update_preview_row_selector(); return
        if not found_placeholders: ttk.Label(self.placeholder_map_frame_content, text=self.lang.get("no_placeholders_found", "No placeholders")).grid(row=0, column=0, columnspan=3, pady=10, sticky="ew"); return
        ttk.Label(self.placeholder_map_frame_content, text=self.lang.get("placeholder_label","Placeholder"), font=("Helvetica", 10, "bold")).grid(row=0, column=0, padx=5, pady=(0,5), sticky="w")
        ttk.Label(self.placeholder_map_frame_content, text=self.lang.get("map_to_excel_column_label","Map to Excel"), font=("Helvetica", 10, "bold")).grid(row=0, column=1, padx=5, pady=(0,5), sticky="w")
        ttk.Label(self.placeholder_map_frame_content, text=self.lang.get("fallback_option_label", "Fallback"), font=("Helvetica", 10, "bold")).grid(row=0, column=2, padx=5, pady=(0,5), sticky="w")
        current_grid_row = 1
        for ph_name in found_placeholders:
            ph_label_widget = ttk.Label(self.placeholder_map_frame_content, text=f"{{{ph_name}}}", anchor="w"); ph_label_widget.grid(row=current_grid_row, column=0, padx=5, pady=2, sticky="ew")
            excel_col_map_var = tk.StringVar(); current_mapping = self.placeholder_config_mappings.get(ph_name)
            if current_mapping and current_mapping in self.excel_column_headers: excel_col_map_var.set(current_mapping)
            elif ph_name.lower() == "name":
                for col_header in self.excel_column_headers:
                    if col_header.lower() == "name": excel_col_map_var.set(col_header); break
            excel_col_combo_widget = ttk.Combobox(self.placeholder_map_frame_content, textvariable=excel_col_map_var, values=[""] + self.excel_column_headers, state="readonly", width=25); excel_col_combo_widget.grid(row=current_grid_row, column=1, padx=5, pady=2, sticky="ew")
            excel_col_combo_widget.bind("<<ComboboxSelected>>", lambda event, p=ph_name, v=excel_col_map_var: self.update_single_placeholder_config_mapping(p, v.get()))
            fallback_text_var = tk.StringVar(); fallback_text_var.set(self.placeholder_fallback_texts.get(ph_name, ""))
            fallback_entry_widget = ttk.Entry(self.placeholder_map_frame_content, textvariable=fallback_text_var, width=30); fallback_entry_widget.grid(row=current_grid_row, column=2, padx=5, pady=2, sticky="ew")
            fallback_text_var.trace_add("write", lambda name, index, mode, p=ph_name, v=fallback_text_var: self.update_single_fallback_text_config(p, v.get()))
            self.placeholder_widgets[ph_name] = {'label_widget': ph_label_widget, 'map_combo_widget': excel_col_combo_widget, 'map_var': excel_col_map_var, 'fallback_entry_widget': fallback_entry_widget, 'fallback_var': fallback_text_var}
            current_grid_row += 1
        self.placeholder_map_frame_content.update_idletasks(); self.placeholder_scrollable_frame.scrollable_area.event_generate("<Configure>"); self.update_preview()

    def update_single_placeholder_config_mapping(self, placeholder_name: str, excel_column_name: str): # Unchanged
        if excel_column_name: self.placeholder_config_mappings[placeholder_name] = excel_column_name
        elif placeholder_name in self.placeholder_config_mappings: del self.placeholder_config_mappings[placeholder_name]
        self.update_preview()

    def update_single_fallback_text_config(self, placeholder_name: str, fallback_text: str): # Unchanged
        self.placeholder_fallback_texts[placeholder_name] = fallback_text; self.update_preview()

    def get_current_placeholder_mappings(self) -> Dict[str, str]: # Unchanged
        current_ui_mappings = {};
        if hasattr(self, 'placeholder_widgets'):
            for ph_name, widget_data_dict in self.placeholder_widgets.items():
                if 'map_var' in widget_data_dict:
                    mapped_col = widget_data_dict['map_var'].get()
                    if mapped_col: current_ui_mappings[ph_name] = mapped_col
        return current_ui_mappings

    def get_current_fallback_texts(self) -> Dict[str, str]: # Unchanged
        current_fallbacks = {};
        if hasattr(self, 'placeholder_widgets'):
            for ph_name, widget_data_dict in self.placeholder_widgets.items():
                if 'fallback_var' in widget_data_dict: current_fallbacks[ph_name] = widget_data_dict['fallback_var'].get()
        return current_fallbacks

    def update_preview_row_selector(self): # Unchanged
        if not (hasattr(self, 'preview_row_combobox') and self.preview_row_combobox.winfo_exists()): print("[UI_UPDATE_SKIP] preview_row_combobox not ready."); return
        if self.loaded_excel_data_preview is not None and not self.loaded_excel_data_preview.empty:
            max_preview_rows = min(len(self.loaded_excel_data_preview), 100)
            preview_choices = [f"Row {i+1}: {' | '.join(str(self.loaded_excel_data_preview.iloc[i, col_idx])[:20] for col_idx in range(min(3, len(self.loaded_excel_data_preview.columns))))}" for i in range(max_preview_rows)]
            self.preview_row_combobox['values'] = preview_choices
            if preview_choices: self.preview_row_combobox.current(0)
            self.preview_row_combobox.config(state="readonly")
        else: self.preview_row_combobox['values'] = []; self.preview_row_var.set(""); self.preview_row_combobox.config(state="disabled")
        self.update_preview()

    # --- MODIFIED: update_preview ---
    def update_preview(self, event=None):
        if not (hasattr(self, 'preview_subject_text') and self.preview_subject_text.winfo_exists() and
                hasattr(self, 'preview_body_text') and self.preview_body_text.winfo_exists()):
            print("[UI_UPDATE_SKIP] Preview widgets not ready.")
            return

        self.preview_subject_text.config(state="normal")
        self.preview_body_text.config(state="normal")
        self.preview_subject_text.delete("1.0", tk.END)
        self.preview_body_text.delete("1.0", tk.END)

        selected_preview_display = self.preview_row_var.get()
        if not selected_preview_display or self.loaded_excel_data_preview is None or self.loaded_excel_data_preview.empty:
            self.preview_subject_text.config(state="disabled")
            self.preview_body_text.config(state="disabled")
            return

        try:
            row_index_match = re.match(r"Row (\d+):", selected_preview_display)
            if not row_index_match:
                self.preview_subject_text.config(state="disabled"); self.preview_body_text.config(state="disabled"); return
            preview_row_index = int(row_index_match.group(1)) - 1
            if not (0 <= preview_row_index < len(self.loaded_excel_data_preview)):
                self.preview_subject_text.config(state="disabled"); self.preview_body_text.config(state="disabled"); return
            row_data = self.loaded_excel_data_preview.iloc[preview_row_index]
        except (ValueError, IndexError):
            self.preview_subject_text.config(state="disabled"); self.preview_body_text.config(state="disabled"); return

        subject_template = self.subject_var.get()
        # Get raw text with user's newlines directly from the body_text_area
        body_template_raw_text = self.body_text_area.get("1.0", tk.END)
        if body_template_raw_text.endswith('\n'): # Remove the single trailing newline tk.Text auto-adds
            body_template_raw_text = body_template_raw_text[:-1]


        active_mappings = self.get_current_placeholder_mappings()
        active_fallbacks = self.get_current_fallback_texts()

        preview_subject_personalized = subject_template
        preview_body_personalized_text = body_template_raw_text # Start with raw text

        for ph_key, excel_col_name in active_mappings.items():
            placeholder_tag = f"{{{ph_key}}}"
            value_to_insert = ""
            if excel_col_name and excel_col_name in row_data:
                cell_value = row_data[excel_col_name]
                if pd.notna(cell_value):
                    value_to_insert = str(cell_value)
                # Ensure blank cell doesn't become "nan" or other string representations of missing
                if pd.isna(cell_value) or (isinstance(cell_value, str) and not cell_value.strip()):
                    value_to_insert = ""


            if not value_to_insert and ph_key in active_fallbacks: # Check if fallback exists
                 value_to_insert = active_fallbacks.get(ph_key, "")


            preview_subject_personalized = preview_subject_personalized.replace(placeholder_tag, value_to_insert)
            preview_body_personalized_text = preview_body_personalized_text.replace(placeholder_tag, value_to_insert)

        # Simple text placeholder for signature in preview
        signature_text_placeholder = ""
        if self.signature_image_path and os.path.exists(self.signature_image_path):
            signature_text_placeholder = f"\n\n[Signature: {os.path.basename(self.signature_image_path)}]"

        final_preview_body_text = preview_body_personalized_text + signature_text_placeholder

        self.preview_subject_text.insert("1.0", preview_subject_personalized)
        self.preview_body_text.insert("1.0", final_preview_body_text)

        self.preview_subject_text.config(state="disabled")
        self.preview_body_text.config(state="disabled")
    # --- END MODIFIED: update_preview ---

    def update_attachment_display(self): # Unchanged
        if not (hasattr(self, 'attachment_var')): return
        if self.attachment_paths: filenames = ", ".join(os.path.basename(path) for path in self.attachment_paths); self.attachment_var.set(filenames[:20] + "..." if len(filenames) > 23 else filenames)
        else: self.attachment_var.set("")

    def select_signature_image(self): # Unchanged
        file_path = filedialog.askopenfilename(title=self.lang.get("signature_button", "Select Signature"), filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif"), ("All files", "*.*")])
        if file_path: self.signature_image_path = file_path; self.update_signature_display(); self.update_preview()

    def clear_signature_image(self): self.signature_image_path = None; self.update_signature_display(); self.update_preview() # Unchanged

    def update_signature_display(self): # Unchanged
        if not (hasattr(self, 'signature_filename_var')): return
        if self.signature_image_path and os.path.exists(self.signature_image_path):
            filename = os.path.basename(self.signature_image_path); display_filename = filename[:8] + "..." if len(filename) > 11 else filename
            self.signature_filename_var.set(self.lang.get("signature_selected_label", "Sig: {filename}").format(filename=display_filename))
        else: self.signature_filename_var.set(self.lang.get("no_signature_selected", "Sig: None"))

    def select_attachments(self): # Unchanged
        file_paths_selected = filedialog.askopenfilenames(title=self.lang.get("select_attach","Select Attachments"), filetypes=[("All files", "*.*")])
        if file_paths_selected:
            for path in file_paths_selected:
                if path and path not in self.attachment_paths: self.attachment_paths.append(path)
        self.update_attachment_display()

    def clear_attachments(self): self.attachment_paths.clear(); self.update_attachment_display() # Unchanged

    def handle_paste(self, event=None): # Unchanged
        try:
            if self.body_text_area.tag_ranges("sel"): self.body_text_area.delete("sel.first", "sel.last")
            clipboard_content = self.root.clipboard_get()
            if "<html" in clipboard_content.lower() or "<body" in clipboard_content.lower():
                 body_match = re.search(r"<body.*?>(.*?)</body>", clipboard_content, re.IGNORECASE | re.DOTALL)
                 if body_match: clipboard_content = body_match.group(1).strip()
            self.body_text_area.insert(tk.INSERT, clipboard_content)
        except tk.TclError: pass
        return "break"

    # --- MODIFIED: get_html_from_text_widget ---
    def get_html_from_text_widget(self):
        if not (hasattr(self, 'body_text_area') and self.body_text_area.winfo_exists()):
            return "<p>&nbsp;</p>" # Should ideally not happen if UI is built

        raw_text_content = self.body_text_area.get("1.0", tk.END)
        # Remove the single trailing newline that tk.Text widget often auto-adds
        if raw_text_content.endswith('\n'):
            raw_text_content = raw_text_content[:-1]

        if not raw_text_content.strip(): # If content is empty or just whitespace after cleanup
            return "<p>&nbsp;</p>"

        lines = raw_text_content.split('\n')
        html_output_parts = []
        current_paragraph_lines = []

        for line_text in lines:
            if not line_text.strip():  # Current line is blank (or only whitespace)
                if current_paragraph_lines:  # Finish the paragraph being built
                    escaped_para_lines = [html.escape(l) for l in current_paragraph_lines]
                    html_output_parts.append(f"<p>{'<br>\n'.join(escaped_para_lines)}</p>")
                    current_paragraph_lines = []
                # Add a representation for the blank line itself
                html_output_parts.append("<p>&nbsp;</p>") # Represent visual blank line as an empty p
            else:
                current_paragraph_lines.append(line_text)
        
        # After loop, if there are any remaining lines for the last paragraph
        if current_paragraph_lines:
            escaped_para_lines = [html.escape(l) for l in current_paragraph_lines]
            html_output_parts.append(f"<p>{'<br>\n'.join(escaped_para_lines)}</p>")
            
        final_html = "\n".join(html_output_parts)
        return final_html if final_html else "<p>&nbsp;</p>" # Ensure something is returned
    # --- END MODIFIED: get_html_from_text_widget ---

    def _apply_text_tag(self, tag_name, **kwargs): # Unchanged
        try:
            if self.body_text_area.winfo_exists():
                self.body_text_area.tag_configure(tag_name, **kwargs)
                if self.body_text_area.tag_ranges("sel"): self.body_text_area.tag_add(tag_name, "sel.first", "sel.last")
        except tk.TclError: pass

    def toggle_bold(self): # Unchanged
        try:
            current_font_obj = font.Font(font=self.body_text_area.cget("font")); new_weight = "bold" if current_font_obj.actual("weight") == "normal" else "normal"
            self._apply_text_tag("user_bold", font=font.Font(family=current_font_obj.actual("family"), size=current_font_obj.actual("size"), weight=new_weight, slant=current_font_obj.actual("slant")))
        except tk.TclError: pass

    def toggle_italic(self): # Unchanged
        try:
            current_font_obj = font.Font(font=self.body_text_area.cget("font")); new_slant = "italic" if current_font_obj.actual("slant") == "roman" else "roman"
            self._apply_text_tag("user_italic", font=font.Font(family=current_font_obj.actual("family"), size=current_font_obj.actual("size"), weight=current_font_obj.actual("weight"), slant=new_slant))
        except tk.TclError: pass

    def apply_font_and_size_change(self, event=None): # Unchanged
        if not (hasattr(self, 'body_text_area') and self.body_text_area.winfo_exists()): return
        try:
            font_size_str = self.font_size_var.get(); font_size = int(font_size_str) if font_size_str.isdigit() and int(font_size_str) > 0 else 14
            self.font_size_var.set(str(font_size)); self.body_text_area.config(font=(self.font_var.get(), font_size)); self.update_preview()
        except ValueError: messagebox.showerror("Error", self.lang.get("invalid_number", "Invalid Number"))
        except tk.TclError as e: logging.warning(f"Font/size apply fail: {e}"); messagebox.showerror("Error", f"Font apply fail: {self.font_var.get()}/{self.font_size_var.get()}")

    def choose_color(self): # Unchanged
        result = colorchooser.askcolor(title=self.lang.get("color_label"))
        if result and result[1]: color_code = result[1]; self._apply_text_tag(f"fg_{color_code.replace('#','')}_{int(time.time())}", foreground=color_code)

    def choose_bg_color(self): # Unchanged
        result = colorchooser.askcolor(title=self.lang.get("bg_color_label"))
        if result and result[1]: color_code = result[1]; self._apply_text_tag(f"bg_{color_code.replace('#','')}_{int(time.time())}", background=color_code)

    def select_excel_file(self): # Unchanged
        file_path = filedialog.askopenfilename(title=self.lang.get("select_excel_prompt", "Select Excel"), filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path: self.excel_file_var.set(file_path); self.loaded_excel_data_preview = load_contacts(file_path, self, self.lang)

    def change_language(self, event=None): # Partially MODIFIED to update new preview_body_label_widget
        print(f"[LANG_CHANGE] Requested. Current: {self.lang_code}, New: {self.language_var.get()}")
        self.lang_code = self.language_var.get(); self.lang = LANGUAGES.get(self.lang_code, LANGUAGES["en"])
        if self.root.winfo_exists(): self.root.title(self.lang["title"])
        print(f"[LANG_CHANGE] Language set to {self.lang_code}. Updating widgets.")

        # Update specific labels that might have been stored as attributes
        if hasattr(self, 'preview_body_label_widget') and self.preview_body_label_widget.winfo_exists():
            self.preview_body_label_widget.config(text=self.lang["preview_body_label"])
            
        # 更新永久激活按钮文本
        if hasattr(self, 'activate_button') and self.activate_button.winfo_exists():
            # 检查当前激活状态
            try:
                status, _ = get_license_status()
                if status == "Activated":
                    self.activate_button.config(text=self.lang["permanently_activated_button"], state=tk.DISABLED)
                else:
                    self.activate_button.config(text=self.lang["permanent_activate_button"], state=tk.NORMAL)
            except:
                # 如果获取状态失败，保持按钮可用
                self.activate_button.config(text=self.lang["permanent_activate_button"], state=tk.NORMAL)

        for widget in self.root.winfo_children(): self._update_widget_language(widget)
        if hasattr(self, 'success_counter_label_widget') and self.success_counter_label_widget.winfo_exists(): self.success_counter_label_widget.config(text=f"{self.lang['success_count_label']} {self.current_success_sends_count}")
        self.update_signature_display()
        if hasattr(self, 'placeholder_mapping_frame') and self.placeholder_mapping_frame.winfo_exists():
            self.placeholder_mapping_frame.config(text=self.lang["personalization_engine_label"])
            if hasattr(self, 'placeholder_scan_instruction_label') and self.placeholder_scan_instruction_label.winfo_exists(): self.placeholder_scan_instruction_label.config(text=self.lang["placeholder_scan_instruction"])
        self.scan_and_setup_placeholder_ui() # This will re-label internal parts of placeholder UI
        if hasattr(self, 'preview_frame') and self.preview_frame.winfo_exists():
            self.preview_frame.config(text=self.lang["preview_engine_label"])
            # The specific preview_body_label_widget is handled above. Others in _update_specific_preview_labels
            for child in self.preview_frame.winfo_children():
                if child != self.preview_body_label_widget: # Avoid double-updating
                    self._update_specific_preview_labels(child)

        if hasattr(self, 'pause_resume_button') and self.pause_resume_button.winfo_exists(): 
            self.pause_resume_button.config(text=self.lang["pause_button"] if self.pause_event.is_set() else self.lang["resume_button"])
            


            
        self.update_preview(); print("[LANG_CHANGE] Widget update complete.")

    def _update_specific_preview_labels(self, widget): # Unchanged (but self.preview_body_label_widget is handled separately now)
        if isinstance(widget, ttk.Frame):
            for grandchild in widget.winfo_children(): self._update_specific_preview_labels(grandchild)
            return
        # Ensure we don't re-translate the one we stored as an attribute
        if hasattr(self, 'preview_body_label_widget') and widget == self.preview_body_label_widget:
            return

        original_texts_map = { LANGUAGES["en"][k]: k for k in ["select_excel_row_for_preview", "refresh_preview_button", "preview_subject_label"]} # Removed preview_body_label
        try:
            current_text = widget.cget("text")
            for text_en_ref, lang_key in original_texts_map.items():
                if any(lang_dict_search.get(lang_key) == current_text for lang_dict_search in LANGUAGES.values()): widget.config(text=self.lang[lang_key]); return
        except (tk.TclError, AttributeError): pass

    def _update_widget_language(self, widget): # Unchanged
        original_text = None
        try:
            if isinstance(widget, (ttk.LabelFrame, ttk.Button, ttk.Label)):
                exempt_widgets = [getattr(self, w_name, None) for w_name in ['success_counter_label_widget', 'signature_label_display', 'placeholder_scan_instruction_label', 'preview_body_label_widget'] if hasattr(self, w_name)]; exempt_widgets.append(getattr(self, 'status_text', None))
                if widget in exempt_widgets or isinstance(widget, ttk.Entry) or (isinstance(widget, tk.Text) and widget != getattr(self, 'status_text', None)): pass
                else: original_text = widget.cget("text")
            if original_text:
                found_key = None
                for lang_key_ref, text_val_en_ref in LANGUAGES["en"].items():
                    if any(lang_dict_search.get(lang_key_ref) == original_text for lang_dict_search in LANGUAGES.values()): found_key = lang_key_ref; break
                if found_key and found_key in self.lang: widget.config(text=self.lang[found_key])
        except (tk.TclError, AttributeError): pass
        if widget != getattr(self, 'placeholder_map_frame_content', None):
            if not (hasattr(self, 'preview_frame') and getattr(widget, 'master', None) == self.preview_frame and isinstance(widget, ttk.Label)):
                 for child in widget.winfo_children(): self._update_widget_language(child)

    def start_sending(self):
        """开始或继续发送邮件"""
        self.status_text.config(state="normal")
        
        # 检查是否有上次保存的进度
        is_continuing = self.current_excel_row > 0
        
        # 如果是全新发送，重置成功计数和当前行索引
        if not is_continuing:
            self.current_success_sends_count = 0
            self.current_excel_row = 0
            self.success_counter_label_widget.config(text=f"{self.lang['success_count_label']} 0")
        else:
            # 如果是继续发送，显示当前进度
            self.log_to_status(f"从第 {self.current_excel_row + 1} 行继续发送 (上次已成功: {self.current_success_sends_count})")
        
        excel_file = self.excel_file_var.get()
        if not excel_file or not os.path.exists(excel_file): 
            messagebox.showerror("Error", self.lang["select_excel_prompt"])
            return
            
        if self.loaded_excel_data_preview is None or self.loaded_excel_data_preview.empty:
            if excel_file: 
                self.loaded_excel_data_preview = load_contacts(excel_file, self, self.lang)
            if self.loaded_excel_data_preview is None: 
                messagebox.showerror("Error", self.lang["load_excel_failed"].format(excel_file=excel_file, error="Data error"))
                return
                
        try:
            # 读取并验证最小/最大延迟参数
            min_delay = float(self.min_delay_var.get())
            max_delay = float(self.max_delay_var.get())
            smtp_port = int(self.smtp_port_var.get())
            batch_size = int(self.batch_size_var.get())
            batch_interval = float(self.batch_interval_var.get())
            
            # 验证数值范围
            if min_delay < 0 or max_delay < 0 or not (0 < smtp_port < 65536) or batch_size <= 0 or batch_interval < 0:
                raise ValueError("Invalid number range")
                
            # 确保最小延迟不大于最大延迟
            if min_delay > max_delay:
                messagebox.showerror("Error", self.lang.get("invalid_delay_range", "Min delay cannot be greater than max delay"))
                return
        except ValueError: 
            messagebox.showerror("Error", self.lang['invalid_number'])
            return
            
        # 检查当前行是否已超出范围
        if self.current_excel_row >= len(self.loaded_excel_data_preview):
            messagebox.showinfo("Info", "所有联系人已处理完毕，将从头开始发送。")
            self.current_excel_row = 0
            self.current_success_sends_count = 0
            self.success_counter_label_widget.config(text=f"{self.lang['success_count_label']} 0")
            
        # 禁用设置控件
        self.toggle_ui_elements_state(tk.DISABLED)
        
        # 更新按钮状态
        self.start_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(state=tk.NORMAL, text=self.lang["pause_button"])
        
        # 设置事件标志
        self.pause_event.set()
        
        # 从当前保存的Excel行开始发送
        start_row = self.current_excel_row
        
        # 传递参数，启动发送线程
        self.sending_thread = threading.Thread(target=self.send_emails_thread, 
                                              args=(min_delay, max_delay, start_row), 
                                              daemon=True)
        self.sending_thread.start()

    def toggle_pause_resume(self): # Unchanged
        if not self.sending_thread or not self.sending_thread.is_alive(): 
            self.pause_resume_button.config(state=tk.DISABLED)
            self.pause_event.set()
            return
            
        if self.pause_event.is_set(): 
            # 暂停操作
            self.pause_event.clear()
            self.pause_resume_button.config(text=self.lang["resume_button"])
            self.log_to_status(self.lang["pausing_sending"])
        else: 
            # 恢复操作
            self.pause_event.set()
            self.pause_resume_button.config(text=self.lang["pause_button"])
            self.log_to_status(self.lang["resuming_sending"])

    def trigger_pause_and_popup(self, error_details: Optional[str]): # Unchanged
        if self.pause_event.is_set():
            self.pause_event.clear()
            if hasattr(self, 'pause_resume_button') and self.pause_resume_button.winfo_exists(): self.pause_resume_button.config(text=self.lang["resume_button"])
            self.log_to_status(self.lang["pausing_sending"] + f" (Auto-paused due to error: {error_details or 'Unknown error'})")
        popup_title = self.lang.get("send_fail_pause_title", "Send Failure"); popup_message_template = self.lang.get("send_fail_pause_message", "Email send failed: {error_details}. Check and resume.")
        display_error = str(error_details or "N/A")[:200]; popup_message = popup_message_template.format(error_details=display_error)
        messagebox.showwarning(popup_title, popup_message)

    def toggle_ui_elements_state(self, state_to_set): # Unchanged
        print(f"[UI_TOGGLE] Setting UI elements to: {state_to_set}")
        frames_to_process_children = []
        if hasattr(self, 'input_frame') and self.input_frame.winfo_exists(): frames_to_process_children.append(self.input_frame)
        if hasattr(self, 'body_text_area') and self.body_text_area.winfo_exists() and self.body_text_area.master.winfo_exists():
            if self.body_text_area.master.winfo_children(): toolbar = self.body_text_area.master.winfo_children()[0]
            if isinstance(toolbar, ttk.Frame): frames_to_process_children.append(toolbar)
        if hasattr(self, 'placeholder_mapping_frame') and self.placeholder_mapping_frame.winfo_exists(): frames_to_process_children.append(self.placeholder_mapping_frame)
        if hasattr(self, 'preview_frame') and self.preview_frame.winfo_exists() and self.preview_frame.winfo_children():
            preview_controls = self.preview_frame.winfo_children()[0]
            if isinstance(preview_controls, ttk.Frame): frames_to_process_children.append(preview_controls)
        if hasattr(self, 'attachment_entry') and self.attachment_entry.winfo_exists() and hasattr(self.attachment_entry.master, 'winfo_children'): frames_to_process_children.append(self.attachment_entry.master)
        for frame_widget in frames_to_process_children:
            if frame_widget and frame_widget.winfo_exists():
                for child in frame_widget.winfo_children():
                    if isinstance(child, (ttk.Entry, ttk.Button, ttk.Combobox)):
                        try: child.config(state=state_to_set)
                        except tk.TclError: pass
        if hasattr(self, 'body_text_area') and self.body_text_area.winfo_exists():
            try: self.body_text_area.config(state=tk.DISABLED if state_to_set == tk.DISABLED else tk.NORMAL)
            except tk.TclError: pass
        if hasattr(self, 'start_button') and self.start_button.winfo_exists():
            try: self.start_button.config(state=state_to_set)
            except tk.TclError: pass
        print(f"[UI_TOGGLE] UI elements state toggling complete for: {state_to_set}")


        
    def email_sending_finished(self):
        """发送完成后的清理操作"""
        self.toggle_ui_elements_state(tk.NORMAL)
        if hasattr(self, 'start_button') and self.start_button.winfo_exists(): 
            self.start_button.config(state=tk.NORMAL)
        if hasattr(self, 'pause_resume_button') and self.pause_resume_button.winfo_exists(): 
            self.pause_resume_button.config(state=tk.DISABLED, text=self.lang["pause_button"])

            
        # 重置状态
        self.pause_event.set()
        self.sending_thread = None

    def send_emails_thread(self, min_delay, max_delay, start_row=0):
        """发送邮件的线程函数，支持从特定行开始发送
        
        Args:
            min_delay: 最小发送延迟（秒）
            max_delay: 最大发送延迟（秒）
            start_row: 开始发送的行号（从0开始）
        """
        sender = self.sender_var.get()
        password = self.password_var.get()
        smtp_host = self.smtp_host_var.get()
        smtp_port_str = self.smtp_port_var.get()
        subject_template = self.subject_var.get()
        cc_email = self.cc_email_var.get() 
        batch_size_str = self.batch_size_var.get()
        batch_interval_str = self.batch_interval_var.get()
        
        # 保存起始行作为当前处理位置
        self.current_excel_row = start_row
        
        try: 
            smtp_port = int(smtp_port_str)
            batch_size = int(batch_size_str)
            batch_interval = float(batch_interval_str)
        except ValueError: 
            self.log_to_status("Error: Invalid numeric setting for sending.")
            self.root.after(0, self.email_sending_finished)
            return
            
        if self.loaded_excel_data_preview is None or self.loaded_excel_data_preview.empty: 
            self.log_to_status(self.lang["no_contacts"])
            self.root.after(0, self.email_sending_finished)
            return
        
        # 检查起始行是否有效
        if start_row >= len(self.loaded_excel_data_preview):
            self.log_to_status(f"起始行 {start_row+1} 超出了数据范围 ({len(self.loaded_excel_data_preview)} 行)")
            self.current_excel_row = 0  # 重置为0，避免下次再次超出范围
            self.root.after(0, self.email_sending_finished)
            return
            
        # 使用整个数据集，但指定起始行
        contacts_df_to_send = self.loaded_excel_data_preview
        body_template_html_raw = self.get_html_from_text_widget()
        active_placeholder_mappings = self.get_current_placeholder_mappings()
        active_placeholder_fallbacks = self.get_current_fallback_texts()
        
        try:
            self.log_to_status(f"Will send to {len(contacts_df_to_send) - start_row} contacts starting from row {start_row+1}. CC: {cc_email}. Delay: {min_delay}-{max_delay}s")
            
            # 将参数传递给发送函数
            final_success_count_for_run = send_emails_to_contacts(
                sender, password, contacts_df_to_send, subject_template, body_template_html_raw,
                active_placeholder_mappings, active_placeholder_fallbacks, self.attachment_paths, cc_email,
                smtp_host, smtp_port, min_delay, max_delay, batch_size, batch_interval, self,
                self.lang, self.success_counter_label_widget, self.pause_event, start_row
            )
            
            # 更新成功计数
            self.current_success_sends_count = final_success_count_for_run
            
            # 如果正常完成，则重置进度
            self.current_excel_row = 0
            summary_msg = self.lang["completed"].format(count=len(contacts_df_to_send) - start_row, cc_email=cc_email)
            try:
                with open("email_summary.txt", "w", encoding="utf-8") as f: 
                    f.write(f"{summary_msg}\n{self.lang['used_excel'].format(excel_file=self.excel_file_var.get())}\n{self.lang['details_log']}\n")
            except Exception as e: 
                logging.error(f"Failed to write summary: {e}")
            self.log_to_status(summary_msg)
        except Exception as e_thread: 
            logging.error(f"Error in sending thread: {e_thread}", exc_info=True)
            self.log_to_status(f"Critical error: {e_thread}")
        finally: 
            self.root.after(0, self.email_sending_finished)

    def save_config(self): # Unchanged
        print("[CONFIG_SAVE] Attempting to save config."); self.placeholder_config_mappings = self.get_current_placeholder_mappings(); self.placeholder_fallback_texts = self.get_current_fallback_texts()
        encrypted_password_for_config = ""; config_password_encryption_key_bytes_str = ""
        if CRYPTO_AVAILABLE:
            try:
                config_password_encryption_key_bytes = Fernet.generate_key(); config_fernet = Fernet(config_password_encryption_key_bytes); password_to_save = self.password_var.get()
                encrypted_password_for_config = config_fernet.encrypt(password_to_save.encode()).decode('utf-8'); config_password_encryption_key_bytes_str = config_password_encryption_key_bytes.decode('utf-8'); print("[CONFIG_SAVE] Password encrypted for saving.")
            except Exception as e_crypt:
                print(f"[CONFIG_SAVE] Password encryption for config failed: {e_crypt}"); logging.error(f"Password encryption: {e_crypt}", exc_info=True)
                if self.root.winfo_exists(): messagebox.showerror("Error", f"无法加密密码: {e_crypt}\n密码将不被保存。"); encrypted_password_for_config = ""; config_password_encryption_key_bytes_str = ""
        else: encrypted_password_for_config = self.password_var.get(); print("[CONFIG_SAVE] Crypto not available. Saving unencrypted.")
        config = {"sender": self.sender_var.get(), "smtp_host": self.smtp_host_var.get(), "smtp_port": self.smtp_port_var.get(), "subject": self.subject_var.get(), "cc_email": self.cc_email_var.get(), "min_delay": self.min_delay_var.get(), "max_delay": self.max_delay_var.get(), "batch_size": self.batch_size_var.get(), "batch_interval": self.batch_interval_var.get(), "excel_file": self.excel_file_var.get(), "attachments": self.attachment_paths, "language": self.language_var.get(), "font_size": self.font_size_var.get(), "font": self.font_var.get(), "signature_image_path": self.signature_image_path, "email_body": self.body_text_area.get("1.0", tk.END).rstrip('\n') if hasattr(self, 'body_text_area') and self.body_text_area.winfo_exists() else self.loaded_body_content, "placeholder_config_mappings": self.placeholder_config_mappings, "placeholder_fallback_texts": self.placeholder_fallback_texts,}
        if CRYPTO_AVAILABLE and config_password_encryption_key_bytes_str: config["password_encrypted"] = encrypted_password_for_config; config["password_encryption_key"] = config_password_encryption_key_bytes_str
        elif not CRYPTO_AVAILABLE and encrypted_password_for_config: config["password"] = encrypted_password_for_config
        try:
            with open(self.config_file, "w", encoding="utf-8") as f: json.dump(config, f, indent=4)
            if self.root.winfo_exists(): messagebox.showinfo("Info", self.lang.get("save_settings_success", "Settings saved successfully")); print("[CONFIG_SAVE] Config saved successfully.")
        except Exception as e:
            print(f"[CONFIG_SAVE] Failed to save config: {e}"); logging.error(f"Failed to save config: {e}", exc_info=True)
            if self.root.winfo_exists(): messagebox.showerror("Error", f"Failed to save settings: {str(e)}")

    def apply_loaded_config_to_ui(self): # Unchanged
        print("[APPLY_UI_CONFIG] Applying loaded config to UI elements..."); self.change_language() # change_language will call update_preview
        loaded_excel_file = self.excel_file_var.get()
        if loaded_excel_file and os.path.exists(loaded_excel_file): print(f"[APPLY_UI_CONFIG] Excel file found: {loaded_excel_file}. Loading..."); self.loaded_excel_data_preview = load_contacts(loaded_excel_file, self, self.lang)
        else: print("[APPLY_UI_CONFIG] No valid Excel file."); self.excel_column_headers = []; self.loaded_excel_data_preview = None
        if hasattr(self, 'body_text_area') and self.body_text_area.winfo_exists():
            if self.loaded_body_content: self.body_text_area.delete("1.0", tk.END); self.body_text_area.insert("1.0", self.loaded_body_content); print("[APPLY_UI_CONFIG] Body content applied.")
        else: print("[APPLY_UI_CONFIG] body_text_area not ready for body content.")
        self.update_attachment_display(); self.apply_font_and_size_change(); self.update_signature_display(); self.scan_and_setup_placeholder_ui(); self.update_preview_row_selector(); # update_preview is called within some of these
        print("[APPLY_UI_CONFIG] UI config application complete.")

    def open_activation_dialog(self):
        """打开激活对话框"""
        print("[MANUAL_ACTIVATION] 用户点击永久激活按钮")
        # 获取当前许可证状态
        try:
            status, data = get_license_status()
            # 显示激活对话框
            self.show_activation_dialog(status, data)
        except Exception as e:
            print(f"[MANUAL_ACTIVATION] 获取许可证状态出错: {e}")
            messagebox.showerror("激活错误", f"无法获取许可证状态: {e}", parent=self.root)

    def update_activate_button_state(self, status="Unknown"):
        """更新永久激活按钮的状态"""
        if hasattr(self, 'activate_button') and self.activate_button.winfo_exists():
            if status == "Activated":
                self.activate_button.config(text=self.lang["permanently_activated_button"], state=tk.DISABLED)
            else:
                self.activate_button.config(text=self.lang["permanent_activate_button"], state=tk.NORMAL)

    # 添加打开帮助网页的函数
    def open_help_website(self):
        """打开帮助网站"""
        print("[HELP] 打开帮助网站")
        try:
            webbrowser.open(self.help_url)
            self.log_to_status(f"已打开帮助网站: {self.help_url}")
        except Exception as e:
            print(f"[HELP] 打开帮助网站失败: {e}")
            messagebox.showerror("错误", f"无法打开帮助网站: {e}", parent=self.root)


if __name__ == "__main__":
    print("[MAIN] Application starting...")
    root = None; app = None
    try:
        root = tk.Tk()
        print("[MAIN] Root Tk window created.")
        if not CRYPTO_AVAILABLE:
            try:
                if root.winfo_exists(): messagebox.showwarning("依赖缺失", "加密库 'cryptography' 未找到或无法导入。\n密码将不会被加密保存，旧的加密密码也无法解密。\n请考虑安装: pip install cryptography", parent=root)
                else: print("WARNING: [MAIN] cryptography missing & root not ready for messagebox.")
            except Exception: print("WARNING: [MAIN] cryptography missing & messagebox failed.")
        
        app = EmailSenderApp(root)
        print("[MAIN] EmailSenderApp initialized. Starting mainloop...")
        if root.winfo_exists(): root.mainloop()
        print("[MAIN] Mainloop exited normally.")
    except Exception as e_main:
        print(f"[MAIN] CRITICAL ERROR in main execution: {e_main}\n{traceback.format_exc()}")
        logging.critical(f"Critical error in Tkinter mainloop: {e_main}", exc_info=True)
        try:
            if root and root.winfo_exists(): messagebox.showerror("致命错误", f"应用程序遇到严重错误并需要关闭。\n详情: {e_main}", parent=root if root.winfo_ismapped() else None)
            else: print(f"致命错误 (UI not available): {e_main}")
        except Exception as e_msgbox: print(f"致命错误 (无法显示消息框): {e_main}\n消息框错误: {e_msgbox}")
    finally:
        print("[MAIN] Application shutting down.")
        if app and hasattr(app, 'sending_thread') and app.sending_thread and app.sending_thread.is_alive(): app.pause_event.set() # Allow thread to attempt graceful exit
        try:
            # 使用更安全的方式检查root是否存在并可被销毁
            if root and hasattr(root, 'destroy') and callable(root.destroy):
                try:
                    # 尝试检查窗口是否存在，但捕获可能的TclError
                    if root.winfo_exists():
                        root.destroy()
                except tk.TclError:
                    # 窗口可能已经被销毁，忽略错误
                    pass
        except Exception as e:
            print(f"[MAIN] Error during final cleanup: {e}")
