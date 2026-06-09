#!/usr/bin/env python3
import os
from pathlib import Path
# 从已经改为 win32com 的 cst_interface 模块中导入
from cst_interface import CSTInterface

def test_cst_reading():
    # 你的工程绝对路径
    project_path = r"D:\workspace\cstwork\original\20260609.cst"
    
    if not os.path.exists(project_path):
        print(f"❌ 错误：在路径 [{project_path}] 找不到该 CST 文件！")
        return

    print("=" * 60)
    print(f"正在启动并连接 CST 进程，打开项目:\n{project_path}")
    print("=" * 60)

    cst = None
    try:
        # 初始化接口，动态加载文件
        cst = CSTInterface(project_path)
        cst.about()

        # 读取仿真的变量信息
        param_names = cst.get_parameter_list()
        total_params = cst.get_number_of_parameters()
        
        print(f"\n✅ 成功连通项目！当前模型内变量总数: {total_params}")
        print("-" * 60)
        print(f"{'变量名 (Parameter)':<25} | {'当前物理值 (Value)':<18} | {'设计表达式 (Expression)'}")
        print("-" * 60)

        for name in param_names:
            try:
                val = cst.get_parameter_value(name)
                expr = cst.get_parameter_expression(name)
                print(f"{name:<25} | {val:<18.4f} | {expr}")
            except Exception as p_err:
                print(f"⚠️ 无法读取变量 [{name}]: {p_err}")

        print("-" * 60)
        print("🎉 测试流程顺利完成。")

    except Exception as e:
        print(f"❌ 自动化控制运行发生错误: {e}")
        
    finally:
        # 无论脚本正常结束还是中间崩溃，都确保关闭 CST 工程连接，防止进程残留
        if cst and cst.project:
            print("\n正在释放工程资源（关闭项目且不保存临时改动）...")
            cst.close_project(save_project=False)
            print("👋 CST 连接已安全释放。")

if __name__ == "__main__":
    test_cst_reading()
