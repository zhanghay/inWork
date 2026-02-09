import os
import shutil

def copy_files_by_txt(txt_path, template_path, output_dir=None):
    """
    根据txt文件每行的前两列内容复制模板文件
    
    参数:
        txt_path: 包含数据的txt文件路径
        template_path: 模板文件tem.xls路径
        output_dir: 输出目录（默认为当前目录）
    """
    # 检查模板文件是否存在
    if not os.path.exists(template_path):
        print(f"❌ 错误: 模板文件 '{template_path}' 不存在！")
        return
    
    # 设置输出目录
    if output_dir is None:
        output_dir = os.path.dirname(txt_path) or '.'
    os.makedirs(output_dir, exist_ok=True)
    
    # 读取并处理txt文件
    try:
        with open(txt_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except FileNotFoundError:
        print(f"❌ 错误: txt文件 '{txt_path}' 不存在！")
        return
    except UnicodeDecodeError:
        # 尝试用gbk编码（中文环境常见）
        try:
            with open(txt_path, 'r', encoding='gbk') as f:
                lines = f.readlines()
        except Exception as e:
            print(f"❌ 错误: 无法读取txt文件（编码问题）: {e}")
            return
    
    success_count = 0
    fail_count = 0
    
    for idx, line in enumerate(lines, 1):
        line = line.strip()
        if not line:  # 跳过空行
            continue
        
        # 尝试多种分隔符：制表符 > 多个空格 > 单个空格 > 逗号
        if '\t' in line:
            parts = line.split('\t')
        elif '  ' in line:  # 多个空格
            parts = line.split()
        elif ',' in line:
            parts = line.split(',')
        else:
            parts = line.split(' ')
        
        # 过滤空字段
        parts = [p.strip() for p in parts if p.strip()]
        
        if len(parts) < 2:
            print(f"⚠️  跳过第 {idx} 行（字段不足2列）: {line}")
            fail_count += 1
            continue
        
        col1, col2 = parts[0], parts[1]
        
        # 清理文件名中的非法字符（Windows）
        invalid_chars = '<>:"/\\|?*'
        clean_col1 = ''.join(c if c not in invalid_chars else '_' for c in col1)
        clean_col2 = ''.join(c if c not in invalid_chars else '_' for c in col2)
        
        new_filename = f"整改单-{clean_col1}-{clean_col2}.xls"
        new_filepath = os.path.join(output_dir, new_filename)
        
        try:
            shutil.copy2(template_path, new_filepath)
            print(f"✅ 已创建: {new_filename}")
            success_count += 1
        except Exception as e:
            print(f"❌ 复制失败（第 {idx} 行）: {new_filename} - {e}")
            fail_count += 1
    
    print(f"\n📊 处理完成: 成功 {success_count} 个, 失败 {fail_count} 个")

# ============ 使用示例 ============
if __name__ == "__main__":
    # 配置路径（请根据实际情况修改）
    TXT_FILE = "title.txt"          # txt数据文件路径
    TEMPLATE_FILE = "temp.xls"      # 模板文件路径
    OUTPUT_DIR = "output"          # 输出目录（可选）
    
    copy_files_by_txt(TXT_FILE, TEMPLATE_FILE, OUTPUT_DIR)