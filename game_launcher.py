import customtkinter as ctk
from tkinter import messagebox, filedialog
import os
import sys
import json
from PIL import Image, ImageDraw
import win32com.client
from icoextract import IconExtractor
from functools import partial  # 引入 partial 确保按钮事件绑定绝对稳定

# ================= 全局配置 =================
CONFIG_FILE = "launcher_config.json"
ICONS_DIR = "icons"

os.makedirs(ICONS_DIR, exist_ok=True)
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")
# ============================================

def get_app_icon(app_path, save_name):
    """尝试从 exe 或 lnk 快捷方式中提取图标"""
    save_path = os.path.join(ICONS_DIR, f"{save_name}.ico")
    
    # 自动去除路径两端可能带有的引号
    app_path = app_path.strip('"').strip("'")
    
    try:
        if app_path.lower().endswith(".lnk"):
            shell = win32com.client.Dispatch("WScript.Shell")
            app_path = shell.CreateShortCut(app_path).Targetpath
        
        if app_path.lower().endswith(".exe"):
            extractor = IconExtractor(app_path)
            extractor.export_icon(save_path)
            return save_path
    except Exception as e:
        print(f"图标提取失败 ({app_path}): {e}")
    return None

def create_fallback_image(text):
    """生成备用纯色占位图标"""
    img = Image.new('RGB', (128, 128), color="#5865F2")
    draw = ImageDraw.Draw(img)
    display_char = text[0].upper() if text else "?"
    draw.text((50, 50), display_char, fill="white")
    return img

class GameLauncherApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gamer 场景启动器")
        
        # 锁定窗口大小并禁用最大化
        self.geometry("700x320")
        self.resizable(False, False) 
        
        self.image_references = []

        # 顶层控制栏
        self.top_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.top_frame.pack(fill="x", pady=(25, 10), padx=30)

        self.title_label = ctk.CTkLabel(self.top_frame, text="🚀 选择你要开启的场景", font=("微软雅黑", 22, "bold"))
        self.title_label.pack(side="left")

        self.add_btn = ctk.CTkButton(self.top_frame, text="➕ 添加新场景", width=120, command=self.open_add_window)
        self.add_btn.pack(side="right")

        # 图标容器
        self.icons_frame = ctk.CTkScrollableFrame(self, orientation="horizontal", fg_color="transparent")
        self.icons_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))

        self.profiles = self.load_config()
        self.render_profiles()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                messagebox.showerror("读取配置失败", f"配置文件损坏: {e}")
        return []

    def save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.profiles, f, ensure_ascii=False, indent=4)

    def render_profiles(self):
        # 清除旧的组件
        for widget in self.icons_frame.winfo_children():
            widget.destroy()
        self.image_references.clear()

        if not self.profiles:
            empty_label = ctk.CTkLabel(self.icons_frame, text="暂无场景，请点击右上角【添加新场景】", text_color="gray", font=("微软雅黑", 14))
            empty_label.pack(pady=50, padx=50)
            return

        for index, profile in enumerate(self.profiles):
            card_frame = ctk.CTkFrame(self.icons_frame, corner_radius=15)
            # anchor="n" 确保卡片紧凑靠上不拉伸
            card_frame.pack(side="left", padx=15, pady=10, anchor="n")

            icon_path = profile.get("icon", "")
            img_ctk = None
            if icon_path and os.path.exists(icon_path):
                try:
                    img = Image.open(icon_path)
                    img_ctk = ctk.CTkImage(light_image=img, dark_image=img, size=(80, 80))
                except:
                    pass
            
            if not img_ctk:
                img = create_fallback_image(profile["name"])
                img_ctk = ctk.CTkImage(light_image=img, dark_image=img, size=(80, 80))
            
            self.image_references.append(img_ctk)

            # 启动按钮 (使用 partial 解决 lambda 在循环中的闭包问题)
            btn = ctk.CTkButton(
                card_frame, image=img_ctk, text="", width=100, height=100, 
                fg_color="transparent", hover_color="#36393F",
                command=partial(self.launch_apps, profile["apps"])
            )
            btn.pack(pady=(15, 5), padx=20)

            name_label = ctk.CTkLabel(card_frame, text=profile["name"], font=("微软雅黑", 14, "bold"))
            name_label.pack(pady=(0, 5))

            edit_btn = ctk.CTkButton(
                card_frame, text="✏️ 编辑", width=70, height=24, 
                fg_color="#4F545C", hover_color="#7289DA", font=("微软雅黑", 11),
                command=partial(self.open_edit_window, index)
            )
            edit_btn.pack(pady=(0, 15))

    def launch_apps(self, apps):
        error_list = []
        for app_path in apps:
            app_path = app_path.strip().strip('"').strip("'")  # 核心修复：清理路径前后的多余空格和引号
            if not app_path: 
                continue
            try:
                os.startfile(app_path)
            except Exception as e:
                error_list.append(f"无法启动: {app_path}\n错误: {e}")
        
        if error_list:
            messagebox.showwarning("部分启动失败", "\n\n".join(error_list))

    def open_add_window(self):
        ProfileWindow(self)

    def open_edit_window(self, index):
        ProfileWindow(self, edit_index=index)


class ProfileWindow(ctk.CTkToplevel):
    def __init__(self, main_app, edit_index=None):
        super().__init__(main_app)
        self.main_app = main_app
        self.edit_index = edit_index 
        self.is_edit_mode = edit_index is not None

        title_text = "编辑场景设定" if self.is_edit_mode else "添加新场景"
        self.title(title_text)
        
        self.geometry("500x560")
        self.resizable(False, False)
        
        # 稳定控制窗口焦点
        self.grab_set() 
        self.attributes("-topmost", True)
        self.after(200, lambda: self.attributes("-topmost", False))

        ctk.CTkLabel(self, text="场景名称:", font=("微软雅黑", 12)).pack(pady=(20, 5), padx=20, anchor="w")
        self.name_entry = ctk.CTkEntry(self, width=460)
        self.name_entry.pack(padx=20)

        ctk.CTkLabel(self, text="图标图片 (留空则自动提取第一个软件的图标):", font=("微软雅黑", 12)).pack(pady=(15, 5), padx=20, anchor="w")
        icon_frame = ctk.CTkFrame(self, fg_color="transparent")
        icon_frame.pack(fill="x", padx=20)
        self.icon_entry = ctk.CTkEntry(icon_frame, width=370)
        self.icon_entry.pack(side="left", padx=(0, 10))
        ctk.CTkButton(icon_frame, text="浏览", width=80, command=self.browse_icon).pack(side="left")

        ctk.CTkLabel(self, text="需要启动的软件路径 (每行一个，可通过调整行来排序):", font=("微软雅黑", 12)).pack(pady=(15, 5), padx=20, anchor="w")
        self.apps_text = ctk.CTkTextbox(self, width=460, height=150)
        self.apps_text.pack(padx=20)
        
        ctk.CTkButton(self, text="➕ 浏览并添加软件(exe/lnk)", fg_color="#4CAF50", hover_color="#45a049", command=self.browse_app).pack(pady=15)

        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.pack(pady=20, fill="x", padx=40)

        if self.is_edit_mode:
            ctk.CTkButton(bottom_frame, text="🗑️ 删除场景", height=40, width=120, fg_color="#ED4245", hover_color="#C82333", font=("微软雅黑", 14, "bold"), command=self.delete_profile).pack(side="left", padx=10)
            ctk.CTkButton(bottom_frame, text="💾 保存修改", height=40, width=120, font=("微软雅黑", 14, "bold"), command=self.save_profile).pack(side="right", padx=10)
            self.load_existing_data() 
        else:
            ctk.CTkButton(bottom_frame, text="💾 保存并应用", height=40, font=("微软雅黑", 14, "bold"), command=self.save_profile).pack()

    def load_existing_data(self):
        profile = self.main_app.profiles[self.edit_index]
        self.name_entry.insert(0, profile.get("name", ""))
        self.icon_entry.insert(0, profile.get("icon", ""))
        apps = profile.get("apps", [])
        self.apps_text.insert("0.0", "\n".join(apps))

    def browse_icon(self):
        path = filedialog.askopenfilename(title="选择图标", filetypes=[("Images", "*.png;*.jpg;*.ico")])
        if path:
            self.icon_entry.delete(0, "end")
            self.icon_entry.insert(0, path)

    def browse_app(self):
        path = filedialog.askopenfilename(title="选择要启动的软件", filetypes=[("Executable/Shortcut", "*.exe;*.lnk"), ("All Files", "*.*")])
        if path:
            current_text = self.apps_text.get("0.0", "end").strip()
            if current_text:
                self.apps_text.insert("end", "\n" + path)
            else:
                self.apps_text.insert("end", path)

    def save_profile(self):
        name = self.name_entry.get().strip()
        icon = self.icon_entry.get().strip()
        apps_raw = self.apps_text.get("0.0", "end").strip().split("\n")
        # 清理用户输入的非法字符和多余引号
        apps = [app.strip().strip('"').strip("'") for app in apps_raw if app.strip()]

        if not name:
            messagebox.showwarning("提示", "请输入场景名称！")
            return
        if not apps:
            messagebox.showwarning("提示", "请至少添加一个需要启动的软件！")
            return

        if not icon:
            first_app = apps[0]
            extracted_icon_path = get_app_icon(first_app, name)
            if extracted_icon_path:
                icon = extracted_icon_path

        new_profile = {
            "name": name,
            "icon": icon,
            "apps": apps
        }

        if self.is_edit_mode:
            self.main_app.profiles[self.edit_index] = new_profile
        else:
            self.main_app.profiles.append(new_profile)

        self.main_app.save_config()
        self.main_app.render_profiles()
        
        self.grab_release() # 释放焦点
        self.destroy()

    def delete_profile(self):
        if messagebox.askyesno("确认删除", f"确定要删除场景「{self.name_entry.get()}」吗？\n删除后不可恢复。"):
            del self.main_app.profiles[self.edit_index]
            self.main_app.save_config()
            self.main_app.render_profiles()
            
            self.grab_release() # 释放焦点
            self.destroy()

if __name__ == "__main__":
    app = GameLauncherApp()
    app.mainloop()