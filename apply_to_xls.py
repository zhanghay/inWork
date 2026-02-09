import os
import glob
import xlwings as xw

def process_files(input_dir="input", output_dir="output"):
    """
    将input文件夹中txt文本内容填入output文件夹同名xls文件
    ✅ 严格遵循：B2/D2/F2 以文本格式填入（防科学计数法）
    ✅ 正确处理 F2 为 F2:J2 合并单元格（仅向F2左上角写入）
    ✅ A5起自动生成序号（1,2,3...），严格对应B列问题行数
    ✅ 完整保留原始xls所有格式（行高/列宽/合并单元格）
    ✅ 不执行任何清空操作（尊重您已手动清除数据）
    """
    os.makedirs(output_dir, exist_ok=True)
    txt_files = glob.glob(os.path.join(input_dir, "*.txt"))
    
    if not txt_files:
        print(f"❌ 未在 {input_dir} 目录下找到任何 .txt 文件")
        return
    
    processed_count = 0
    for txt_path in txt_files:
        filename = os.path.splitext(os.path.basename(txt_path))[0]
        xls_path = os.path.join(output_dir, f"{filename}.xls")
        
        if not os.path.exists(xls_path):
            print(f"⚠️  跳过 {filename}：未找到对应的Excel文件 {xls_path}")
            continue
        
        try:
            # 读取并解析txt内容
            with open(txt_path, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f.readlines() if line.strip()]
            
            if not lines:
                print(f"⚠️  跳过 {filename}：txt文件为空")
                continue
            
            # 解析第一行：工单编号-户号-户名（仅分割前两个'-'）
            parts = lines[0].split('-', 2)
            if len(parts) < 3:
                print(f"⚠️  跳过 {filename}：第一行格式错误（需至少包含两个'-'）")
                continue
            
            gongdan_id, hu_hao, hu_ming = [p.strip() for p in parts]
            
            # 从第三行开始提取问题描述（严格遵循需求：索引2 = 第三行）
            issue_lines = lines[2:] if len(lines) > 2 else []
            
            # 使用xlwings操作Excel（后台模式）
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            
            try:
                wb = app.books.open(xls_path)
                sht = wb.sheets[0]
                
                # === 步骤1：设置B2/D2/F2为文本格式并填入 ===
                # 关键：F2是F2:J2合并单元格，仅向左上角F2写入即可
                for cell_addr, value in [('B2', gongdan_id), ('D2', hu_hao), ('F2', hu_ming)]:
                    cell = sht.range(cell_addr)
                    cell.number_format = '@'  # 强制文本格式
                    cell.value = value
                
                # === 步骤2：填入问题描述（B5起，B:C:D为合并单元格）===
                # 注意：不清空原有数据（尊重您已手动清除）
                start_row = 5
                for idx, issue in enumerate(issue_lines):
                    sht.range(f'B{start_row + idx}').value = issue
                
                # === 步骤3：生成序号（A5起，仅对实际填入的问题行生成）===
                if issue_lines:
                    # 生成1到n的序号（列向量格式 [[1],[2],[3]]）
                    seq_numbers = [[i + 1] for i in range(len(issue_lines))]
                    sht.range(f'A{start_row}').value = seq_numbers
                else:
                    # 无问题描述时，确保A5为空（避免残留旧序号）
                    sht.range('A5').value = None
                
                # 保存并关闭
                wb.save()
                wb.close()
                processed_count += 1
                print(f"✅ 成功处理: {filename} | 工单:{gongdan_id} | 问题项:{len(issue_lines)}")
                
            finally:
                app.quit()
                
        except Exception as e:
            print(f"❌ 处理 {filename} 时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            if 'app' in locals():
                try:
                    app.quit()
                except:
                    pass
    
    print(f"\n📊 处理完成: {processed_count}/{len(txt_files)} 个文件成功处理")

if __name__ == "__main__":
    process_files()