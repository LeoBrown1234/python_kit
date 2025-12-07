"""
批量解压文件工具，交互式输入。
在输入指定文件夹后，输入统一压缩密码，进行压缩。
优点是界面美观。
未来有打包文件的准备。
"""
import os
import sys
import pyzipper  # 替代zipfile，支持AES加密
from tqdm import tqdm


def get_valid_directory():
    """交互获取有效目录路径，带提示+输入结果显示"""
    while True:
        print("\n===== 请输入要压缩的目录路径 =====")
        dir_path = input("👉 目录路径（支持中文，例如：D:\\我的文件）：").strip()
        dir_path = os.path.expanduser(dir_path)
        print(f"📌 你输入的目录路径：{dir_path}")

        if os.path.isdir(dir_path):
            abs_dir = os.path.abspath(dir_path)
            print(f"✅ 目录验证通过，实际使用路径：{abs_dir}")
            return abs_dir
        else:
            print(f"❌ 错误：路径「{dir_path}」不是有效目录，请重新输入！")


def get_compress_password():
    """交互获取压缩密码（带提示+输入结果显示，密码显示为*）"""
    while True:
        print("\n===== 请设置压缩密码（可选） =====")
        print("提示：直接按回车表示不设置密码")
        pwd1 = input("👉 请输入压缩密码：").strip()
        pwd_display = "*" * len(pwd1) if pwd1 else "（无密码）"
        print(f"📌 你输入的密码：{pwd_display}")

        if not pwd1:
            print("✅ 确认不设置压缩密码")
            return None

        print("\n===== 请确认压缩密码 =====")
        pwd2 = input("👉 请再次输入密码确认：").strip()
        pwd2_display = "*" * len(pwd2) if pwd2 else "（无密码）"
        print(f"📌 你再次输入的密码：{pwd2_display}")

        if pwd1 == pwd2:
            print("✅ 两次密码输入一致，密码设置完成")
            return pwd1
        else:
            print("❌ 两次输入的密码不一致，请重新输入！")


def compress_single_item(item_path, zip_path, password=None):
    """
    用pyzipper压缩单个文件/文件夹（支持AES加密+中文+进度）
    :param item_path: 待压缩的文件/文件夹绝对路径
    :param zip_path: 输出ZIP文件的绝对路径
    :param password: 压缩密码（None表示无密码）
    """
    # 收集所有待压缩的文件路径
    file_list = []
    if os.path.isfile(item_path):
        file_list.append(item_path)
    else:
        for root, _, files in os.walk(item_path):
            for file in files:
                file_list.append(os.path.join(root, file))

    # 计算总文件大小（用于进度条）
    total_size = 0
    valid_files = []
    for f in file_list:
        try:
            total_size += os.path.getsize(f)
            valid_files.append(f)
        except Exception as e:
            print(f"\n⚠️  警告：无法获取文件「{f}」大小 - {e}，跳过该文件")

    if not valid_files:
        print(f"\nℹ️  提示：「{os.path.basename(item_path)}」内无有效文件，跳过压缩")
        return

    # 初始化ZIP文件（AES-256加密，支持中文）
    with pyzipper.AESZipFile(
            zip_path,
            'w',
            compression=pyzipper.ZIP_DEFLATED,  # 压缩模式
            encryption=pyzipper.WZ_AES if password else None  # 有密码则用AES加密
    ) as zipf:
        # 设置密码（AES-256）
        if password:
            zipf.setpassword(password.encode('utf-8'))

        # 进度条
        with tqdm(
                total=total_size,
                unit='B',
                unit_scale=True,
                unit_divisor=1024,
                desc=f"🚀 正在压缩「{os.path.basename(item_path)}」"
        ) as pbar:
            for file_path in valid_files:
                try:
                    # 保持原目录结构（解决中文路径）
                    rel_path = os.path.relpath(file_path, os.path.dirname(item_path))
                    zipf.write(file_path, rel_path)

                    # 更新进度
                    file_size = os.path.getsize(file_path)
                    pbar.update(file_size)
                except Exception as e:
                    print(f"\n⚠️  警告：压缩文件「{file_path}」失败 - {e}，跳过该文件")


def batch_compress():
    """批量压缩目录内所有文件/文件夹"""
    print("=" * 50)
    print("📁 目录文件批量压缩工具（纯Python ZIP加密版）")
    print("=" * 50)

    # 获取基础信息
    target_dir = get_valid_directory()
    password = get_compress_password()

    # 获取待处理项
    items = [os.path.join(target_dir, item) for item in os.listdir(target_dir)]
    print(f"\n===== 压缩任务初始化 =====")
    print(f"📂 待压缩目录：{target_dir}")
    print(f"🔑 密码设置状态：{'已设置（AES-256加密）' if password else '未设置'}")
    print(f"📊 待压缩项目总数：{len(items)}")

    if not items:
        print("ℹ️  提示：指定目录内无任何文件/文件夹，无需压缩！")
        return

    # 逐个压缩
    success_count = 0
    print(f"\n===== 开始批量压缩 =====\n")
    for idx, item in enumerate(items, 1):
        item_name = os.path.basename(item)
        zip_name = f"{item_name}.zip"
        zip_path = os.path.join(target_dir, zip_name)

        print(f"\n[{idx}/{len(items)}] 处理项：{item_name}")

        # 避免覆盖已存在的ZIP文件
        if os.path.exists(zip_path):
            print(f"⚠️  「{zip_name}」已存在，跳过压缩")
            continue

        try:
            compress_single_item(item, zip_path, password)
            success_count += 1
            print(f"✅ 「{zip_name}」压缩完成 → 保存路径：{zip_path}")
        except Exception as e:
            print(f"❌ 「{zip_name}」压缩失败 - {e}")
            # 删除失败的不完整ZIP文件
            if os.path.exists(zip_path):
                os.remove(zip_path)
                print(f"🗑️  已删除不完整的压缩文件：{zip_path}")

    # 输出汇总信息
    print(f"\n" + "=" * 50)
    print(f"📋 压缩任务汇总")
    print(f"=" * 50)
    print(f"📂 压缩目录：{target_dir}")
    print(f"🔑 密码状态：{'已设置（AES-256加密）' if password else '未设置'}")
    print(f"📊 总项目数：{len(items)}")
    print(f"✅ 成功数：{success_count}")
    print(f"❌ 失败/跳过数：{len(items) - success_count}")
    print(f"💾 所有压缩文件均保存在：{target_dir}")


if __name__ == "__main__":
    # 设置系统编码为UTF-8，解决中文显示问题
    if sys.platform == "win32":
        os.system("chcp 65001 > nul")
    sys.stdout.reconfigure(encoding='utf-8')

    try:
        batch_compress()
    except KeyboardInterrupt:
        print("\n\n🛑 用户中断操作，程序退出！")
    except Exception as e:
        print(f"\n❌ 程序运行出错：{e}")
    finally:
        input("\n\n按回车键退出程序...")