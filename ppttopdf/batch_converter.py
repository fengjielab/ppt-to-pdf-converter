import os
import sys
import comtypes.client

def batch_ppt_to_pdf(input_folder):
    """
    批量将指定文件夹内的所有 PPT 和 PPTX 文件转换为 PDF。

    :param input_folder: 包含 PowerPoint 文件的文件夹路径。
    """
    # 检查输入文件夹是否存在
    if not os.path.isdir(input_folder):
        print(f"错误: 文件夹 '{input_folder}' 不存在。")
        return

    # 获取文件夹的绝对路径
    input_folder = os.path.abspath(input_folder)
    
    powerpoint = None
    try:
        # 启动 PowerPoint 应用程序，并设置为不可见
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 0  # 0 for False, 1 for True. Keep it in the background.

        print(f"开始扫描文件夹: {input_folder}")
        
        # 遍历文件夹中的所有文件
        for filename in os.listdir(input_folder):
            # 检查文件是否为 PowerPoint 文件
            if filename.lower().endswith(('.ppt', '.pptx')):
                
                # 构建完整的输入文件路径
                input_file_path = os.path.join(input_folder, filename)
                
                # 构建输出的 PDF 文件路径（与原文件同名，扩展名不同）
                output_file_path = os.path.splitext(input_file_path)[0] + ".pdf"
                
                print(f"  -> 正在转换: {filename} ...")
                
                presentation = None # 确保变量存在
                try:
                    # 打开演示文稿
                    presentation = powerpoint.Presentations.Open(input_file_path, WithWindow=False)
                    
                    # 另存为 PDF (formatType=32)
                    presentation.SaveAs(output_file_path, 32)
                    print(f"  ✔ 成功转换为: {os.path.basename(output_file_path)}")
                    
                except Exception as e:
                    print(f"  ❌ 转换文件 '{filename}' 时发生错误: {e}")
                finally:
                    # 确保关闭已打开的演示文稿
                    if presentation:
                        presentation.Close()

    except Exception as e:
        print(f"处理过程中发生严重错误: {e}")
        print("请确保您的电脑已正确安装 Microsoft PowerPoint。")
    finally:
        # 确保 PowerPoint 应用程序在所有操作完成后退出
        if powerpoint:
            powerpoint.Quit()
            print("\n所有转换任务完成，PowerPoint 已退出。")

# --- 使用示例 ---
if __name__ == '__main__':
    # 检查是否在 Windows 上运行
    if sys.platform != 'win2d32':
        print("错误：此脚本需要 Windows 操作系统和 Microsoft PowerPoint。")
    else:
        # --- 请在这里设置您的 PPT 文件夹路径 ---
        # 使用 '.' 表示当前文件夹 (即脚本所在的文件夹)
        # 或者指定一个绝对路径, 例如: 'C:\\Users\\YourName\\Documents\\MyPresentations'
        folder_to_process = '.' 
        
        batch_ppt_to_pdf(folder_to_process)