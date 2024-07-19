import pandas as pd
import numpy as np
from sklearn.metrics import pairwise_distances
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageTk
import subprocess

# 读取 CSV 文件
data = pd.read_csv('datacollect071802.csv')

# 排除不需要显示和导出的列
columns_to_exclude = ['C', 'h']
data = data.drop(columns=columns_to_exclude, errors='ignore')

# 打印列名以确认
print(data.columns)

# 清理数据中的 NaN 值
data.dropna(subset=['L', 'A', 'B'], inplace=True)

# 计算 delta E 值
def calculate_delta_e(lab_value, database):
    lab_values = database[['L', 'A', 'B']].values
    delta_e = pairwise_distances(lab_values, [lab_value], metric='euclidean')
    return delta_e.flatten()

# 查找最接近的配方，避免重复的 Serial Number
def find_closest_recipes(lab_value, database, include_word=None, exclude_word=None, top_n=3):
    delta_e = calculate_delta_e(lab_value, database)
    database['Delta E'] = delta_e
    database = database.sort_values(by='Delta E')
    
    unique_serials = []
    closest_recipes = []
    for index, row in database.iterrows():
        if row['Serial Number'] not in unique_serials:
            recipe_str = ' '.join([str(row[col]) for col in ['砂粉料一', '砂粉料二', '砂粉料三', '砂粉料四', '砂粉料五'] if col in row])
            if (include_word and include_word not in recipe_str) or (exclude_word and exclude_word in recipe_str):
                continue
            unique_serials.append(row['Serial Number'])
            closest_recipes.append(row)
        if len(closest_recipes) == top_n:
            break
    
    closest_recipes_df = pd.DataFrame(closest_recipes)
    if closest_recipes_df.empty:
        return closest_recipes_df, np.array([])
    return closest_recipes_df, closest_recipes_df['Delta E'].values

def describe_color_difference(input_lab, match_lab):
    descriptions = []
    labels = ['L', 'A', 'B']
    for i, label in enumerate(labels):
        if match_lab[i] > input_lab[i]:
            if label == 'L':
                descriptions.append("此配方需要 暗一點或少一點白")
            elif label == 'A':
                descriptions.append("绿一點或少一點红")
            elif label == 'B':
                descriptions.append("藍一點或少一點黄")
        else:
            if label == 'L':
                descriptions.append("白一點或少一點黑")
            elif label == 'A':
                descriptions.append("红一點或少一點綠")
            elif label == 'B':
                descriptions.append("黄一點或少一點藍")
    return '，'.join(descriptions)

# 使用 xicclu 进行 LAB 转 CMYK 转换
def lab_to_cmyk_with_icc(lab_value, icc_profile_path):
    lab_str = f"{lab_value[0]} {lab_value[1]} {lab_value[2]}"
    try:
        # 使用 xicclu 进行颜色转换
        process = subprocess.Popen(
            ['/opt/homebrew/bin/xicclu', '-fb', '-ir', '-pl', icc_profile_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        stdout, stderr = process.communicate(input=lab_str)
        print(f"Command output:\n{stdout}")
        print(f"Command error output:\n{stderr}")
        if process.returncode == 0:
            cmyk_value = extract_cmyk_value(stdout)
            if cmyk_value:
                print(f"CMYK: {cmyk_value}")
                return cmyk_value
            else:
                print(f"No valid CMYK value found in output.")
                return None
        else:
            print(f"Error: {stderr}")
            return None
    except subprocess.CalledProcessError as e:
        print(f"Command '{e.cmd}' returned non-zero exit status {e.returncode}.")
        print(f"Error output: {e.stderr}")
        return None
    except Exception as e:
        print(f"Exception during conversion: {e}")
        return None

def extract_cmyk_value(xicclu_output):
    lines = xicclu_output.strip().split("\n")
    for line in lines:
        if "->" in line and "[CMYK]" in line:
            parts = line.split(" ")
            cmyk_index = parts.index('[CMYK]')
            cmyk_values = parts[cmyk_index - 4:cmyk_index]
            return cmyk_values
    return None

# 导出 CSV 文件
def export_to_csv(recipes, filename="exported_recipes.csv"):
    recipes.to_csv(filename, index=False)
    print(f"Recipes exported to {filename}")

# GUI 界面
class LabMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("LAB Value Matcher")
        self.root.geometry("1000x1080")  # 设置窗口大小
        self.root.configure(bg='black')  # 设置背景色为黑色
        
        self.icc_profile_path = None
        self.lab_value = [tk.DoubleVar() for _ in range(3)]
        
        frame = ttk.Frame(root, padding="10", style='My.TFrame')
        frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)  # 将frame置于窗口中央
        
        # 添加 logo
        self.logo = Image.open('qjlogo.jpg')
        self.logo = self.logo.resize((588, 173), Image.Resampling.LANCZOS)  # 按比例缩小
        self.logo = ImageTk.PhotoImage(self.logo)
        logo_label = ttk.Label(frame, image=self.logo, background='black')
        logo_label.grid(column=0, row=0, columnspan=2, pady=10)
        
        labels = ["測量L值", "測量a值", "測量b值"]
        for i, label in enumerate(labels):
            ttk.Label(frame, text=label,font=('Helvetica', 15, 'bold'), background='black', foreground='white').grid(column=0, row=i+1, padx=5, pady=5)
            ttk.Entry(frame, textvariable=self.lab_value[i]).grid(column=1, row=i+1, padx=5, pady=5)
        
        self.include_word = tk.StringVar()
        self.exclude_word = tk.StringVar()
        
        ttk.Label(frame, text="配方砂粉指定", font=('Helvetica', 15, 'bold'),background='black', foreground='white').grid(column=0, row=4, padx=5, pady=5)
        ttk.Entry(frame, textvariable=self.include_word).grid(column=1, row=4, padx=5, pady=5)
        
        ttk.Label(frame, text="配方砂粉排除", font=('Helvetica', 15, 'bold'),background='black', foreground='white').grid(column=0, row=5, padx=5, pady=5)
        ttk.Entry(frame, textvariable=self.exclude_word).grid(column=1, row=5, padx=5, pady=5)
        
        style = ttk.Style()
        style.configure('TButton', background='black', foreground='white', width=20)
        style.configure('My.TFrame', background='black')

        match_button = ttk.Button(frame, text="開始配對", command=self.match_lab_values, style='TButton')
        match_button.grid(column=0, row=6, padx=5, pady=5)

        load_icc_button = ttk.Button(frame, text="導入 ICC Profile", command=self.load_icc_profile, style='TButton')
        load_icc_button.grid(column=1, row=6, padx=5, pady=5)

        # Notebook 用于显示配方结果
        self.notebook = ttk.Notebook(frame)
        self.notebook.grid(column=0, row=7, columnspan=2, padx=5, pady=5)

        self.recipe_tabs = []
        self.result_texts = []
        self.delta_e_labels = []
        self.cmyk_labels = []
        self.warning_labels = []
        for i in range(3):
            tab = ttk.Frame(self.notebook, style='My.TFrame')
            self.notebook.add(tab, text=f'最接近配方 {i+1}')
            self.recipe_tabs.append(tab)

            warning_label = ttk.Label(tab, text="", font=('Helvetica', 12, 'bold'), background='black', foreground='white')
            warning_label.grid(column=0, row=0, padx=5, pady=5)
            self.warning_labels.append(warning_label)

            delta_e_label = ttk.Label(tab, text=f"Delta E: N/A", font=('Helvetica', 14, 'bold'), background='black', foreground='white')
            delta_e_label.grid(column=0, row=1, padx=5, pady=5)
            self.delta_e_labels.append(delta_e_label)

            result_frame = ttk.Frame(tab, style='My.TFrame')
            result_frame.grid(column=0, row=2, padx=5, pady=5)
            self.result_texts.append(result_frame)

            cmyk_label = ttk.Label(tab, text=f"CMYK: N/A", font=('Helvetica', 14, 'bold'), background='black', foreground='white')
            cmyk_label.grid(column=0, row=3, padx=5, pady=5)
            self.cmyk_labels.append(cmyk_label)

        export_button = ttk.Button(frame, text="匯出配方資料", command=self.export_recipes, style='TButton')
        export_button.grid(column=0, row=8, columnspan=2, padx=5, pady=5)

    def update_lab_values(self, lab_values):
        for i, value in enumerate(lab_values):
            self.lab_value[i].set(value)
        self.match_lab_values()

    def match_lab_values(self):
        lab_value = [var.get() for var in self.lab_value]
        if any(np.isnan(lab_value)):
            print("Input contains NaN.")
            return

        include_word = self.include_word.get().strip() or None
        exclude_word = self.exclude_word.get().strip() or None

        # 清除之前的结果
        self.warning_labels.clear()
        self.delta_e_labels.clear()
        self.result_texts.clear()
        self.cmyk_labels.clear()
        
        for tab in self.recipe_tabs:
            for widget in tab.winfo_children():
                widget.destroy()

        closest_recipes, delta_e_values = find_closest_recipes(lab_value, data, include_word, exclude_word)

        if closest_recipes.empty or delta_e_values.size == 0:
            for tab in self.recipe_tabs:
                ttk.Label(tab, text="查無合適配方", font=('Helvetica', 14, 'bold'), background='black', foreground='red').grid(column=0, row=0, padx=5, pady=5)
            return

        for i, (index, row), delta_e in zip(range(3), closest_recipes.iterrows(), delta_e_values):
            warning_label = ttk.Label(self.recipe_tabs[i], text="", font=('Helvetica', 12, 'bold'), background='black', foreground='white')
            warning_label.grid(column=0, row=0, padx=5, pady=5)
            self.warning_labels.append(warning_label)

            delta_e_label = ttk.Label(self.recipe_tabs[i], text=f"Delta E: N/A", font=('Helvetica', 14, 'bold'), background='black', foreground='white')
            delta_e_label.grid(column=0, row=1, padx=5, pady=5)
            self.delta_e_labels.append(delta_e_label)

            result_frame = ttk.Frame(self.recipe_tabs[i], style='My.TFrame')
            result_frame.grid(column=0, row=2, padx=5, pady=5)
            self.result_texts.append(result_frame)

            cmyk_label = ttk.Label(self.recipe_tabs[i], text=f"CMYK: N/A", font=('Helvetica', 14, 'bold'), background='black', foreground='white')
            cmyk_label.grid(column=0, row=3, padx=5, pady=5)
            self.cmyk_labels.append(cmyk_label)

            if delta_e > 1:
                warning_label.config(text="Delta E過大 配方不準確", foreground='red')
            else:
                warning_label.config(text="Delta E 合理範圍，配方具參考性", foreground='green')

            delta_e_label.config(text=f"Delta E: {delta_e:.2f}")
            color_diff_desc = describe_color_difference(lab_value, row[['L', 'A', 'B']].values)

            # 清空之前的内容
            for widget in result_frame.winfo_children():
                widget.destroy()

            # 动态生成显示区域
            row_data = row.dropna()  # 去除 NaN 值
            for j, (col, val) in enumerate(row_data.items()):
                if col not in ['L', 'A', 'B', 'Delta E']:
                    ttk.Label(result_frame, text=f"{col}:", font=('Helvetica', 14, 'bold'), background='black', foreground='white').grid(column=0, row=j, padx=5, pady=5, sticky=tk.W)
                    ttk.Label(result_frame, text=f"{val}", font=('Helvetica', 14), background='black', foreground='white').grid(column=1, row=j, padx=5, pady=5, sticky=tk.W)

            ttk.Label(result_frame, text=f"\n顏色差異描述: {color_diff_desc}", font=('Helvetica', 14), background='black', foreground='white').grid(column=0, row=len(row_data), columnspan=2, padx=5, pady=5, sticky=tk.W)

            if self.icc_profile_path:
                cmyk_value = lab_to_cmyk_with_icc(lab_value, self.icc_profile_path)
                if cmyk_value:
                    cmyk_text = f"C: {float(cmyk_value[0]) * 100:.0f}%, M: {float(cmyk_value[1]) * 100:.0f}%, Y: {float(cmyk_value[2]) * 100:.0f}%, K: {float(cmyk_value[3]) * 100:.0f}%"
                    cmyk_label.config(text=f"CMYK: {cmyk_text}")
                else:
                    cmyk_label.config(text="CMYK: Conversion failed")

    def load_icc_profile(self):
        self.icc_profile_path = filedialog.askopenfilename(filetypes=[("ICC Profiles", "*.icc *.icm")])
        if self.icc_profile_path:
            print(f"Loaded ICC profile: {self.icc_profile_path}")

    def export_recipes(self):
        lab_value = [var.get() for var in self.lab_value]
        include_word = self.include_word.get().strip() or None
        exclude_word = self.exclude_word.get().strip() or None
        closest_recipes, _ = find_closest_recipes(lab_value, data, include_word, exclude_word)
        # 去除不需要的列再导出
        closest_recipes = closest_recipes.drop(columns=columns_to_exclude, errors='ignore')
        export_to_csv(closest_recipes)

# 运行应用
if __name__ == "__main__":
    root = tk.Tk()
    app = LabMatcherApp(root)
    root.mainloop()
