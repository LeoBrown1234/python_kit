"""
批量解压文件工具，交互式输入。
在输入指定文件夹后，输入统一解压密码，进行解压。
优点是界面美观。
未来有打包文件的准备。
"""
# -*- coding: utf-8 -*-
import os
import sys
import zipfile
import rarfile
import py7zr
import time
from pathlib import Path
from dataclasses import dataclass


# 数据类：存储进度信息
@dataclass
class ExtractProgress:
    total_files: int  # 总文件数
    current_index: int  # 当前解压文件索引
    remaining_files: int  # 剩余文件数
    total_size: int  # 所有压缩文件总大小(字节)
    current_file_size: int  # 当前文件大小(字节)
    current_file_unpacked: int  # 当前文件已解压大小(字节)
    remaining_size: int  # 剩余文件总大小(字节)


# 字节大小格式化（转换为KB/MB/GB）
def format_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


# 实时进度显示函数
def print_progress(progress: ExtractProgress, current_file_name: str):
    # 计算当前文件解压百分比
    current_percent = (progress.current_file_unpacked / progress.current_file_size * 100
                       ) if progress.current_file_size > 0 else 0

    # 计算整体进度百分比
    overall_percent = ((progress.total_files - progress.remaining_files) / progress.total_files * 100
                       ) if progress.total_files > 0 else 0

    # 构建进度信息
    progress_info = f"""
┌─────────────────────────────────────────────────────┐
│  整体进度：{progress.current_index}/{progress.total_files} ({overall_percent:.1f}%)  剩余文件：{progress.remaining_files}个  │
│  当前解压：{current_file_name} ({current_percent:.1f}%)                          │
│  文件大小：{format_size(progress.current_file_size)}  已解压：{format_size(progress.current_file_unpacked)}     │
│  剩余总大小：{format_size(progress.remaining_size)}                              │
└─────────────────────────────────────────────────────┘
"""
    # 清屏并打印进度（Windows/Linux兼容）
    os.system('cls' if os.name == 'nt' else 'clear')
    sys.stdout.write(progress_info)
    sys.stdout.flush()


# 自定义ZIP解压进度回调
class ZipExtractProgress:
    def __init__(self, total_size: int, progress_obj: ExtractProgress):
        self.total_size = total_size
        self.progress_obj = progress_obj
        self.unpacked = 0

    def update(self, bytes_amount: int):
        self.unpacked += bytes_amount
        self.progress_obj.current_file_unpacked = self.unpacked
        print_progress(self.progress_obj, Path(self.progress_obj.current_file_name).name)
        time.sleep(0.01)  # 避免刷新过快


# 自定义RAR/7Z解压进度跟踪（通过文件分块模拟）
def track_extract_progress(file_path: str, extract_func, progress_obj: ExtractProgress):
    """通用进度跟踪包装器"""
    file_size = os.path.getsize(file_path)
    progress_obj.current_file_size = file_size
    progress_obj.current_file_unpacked = 0

    # 模拟进度更新（实际解压无法精确获取字节级进度，按文件数/分块估算）
    def progress_hook(unpacked: int):
        progress_obj.current_file_unpacked = unpacked
        print_progress(progress_obj, Path(file_path).name)

    # 执行解压并跟踪进度
    extract_func(progress_hook)
    # 最终刷新进度为100%
    progress_obj.current_file_unpacked = file_size
    print_progress(progress_obj, Path(file_path).name)


def validate_directory(path: str) -> Path:
    """验证目录是否存在，不存在则创建（适配中文路径）"""
    dir_path = Path(path).absolute()
    if not dir_path.exists():
        dir_path.mkdir(parents=True, exist_ok=True)
        print(f"目录不存在，已创建：{dir_path}")
    if not dir_path.is_dir():
        raise ValueError(f"路径 {dir_path} 不是有效的目录！")
    return dir_path


def extract_zip(file_path: Path, extract_dir: Path, password: str = None, progress_obj: ExtractProgress = None):
    """解压 zip 文件（适配中文路径+进度显示）"""
    try:
        zipfile._get_decompressors = lambda: None
        file_path_str = str(file_path.absolute())
        target_dir = extract_dir / file_path.stem
        target_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(file_path_str, 'r') as zf:
            # 获取ZIP文件总大小
            total_size = sum(file.file_size for file in zf.infolist())
            progress_obj.current_file_name = file_path_str
            progress_obj.current_file_size = total_size

            # 设置密码
            if password:
                zf.setpassword(password.encode('utf-8'))

            # 自定义进度回调
            progress = ZipExtractProgress(total_size, progress_obj)

            # 逐个解压文件并更新进度
            for file in zf.infolist():
                try:
                    # 处理中文文件名
                    filename = file.filename
                    try:
                        filename = filename.encode('cp437').decode('gbk')
                    except:
                        filename = filename.encode('utf-8').decode('utf-8')

                    # 解压文件并更新进度
                    zf.extract(file, target_dir, pwd=password.encode('utf-8') if password else None)
                    progress.update(file.file_size)
                except Exception as e:
                    print(f"\n⚠️  解压文件 {file.filename} 时出错：{e}")
                    continue

        print(f"\n✅ 解压完成：{file_path.name} -> {target_dir}")
        return True
    except zipfile.BadZipFile:
        print(f"\n❌ 错误：{file_path.name} 不是有效的 ZIP 文件")
        return False
    except RuntimeError as e:
        if "password" in str(e).lower():
            print(f"\n❌ 错误：{file_path.name} 解压密码错误")
        else:
            print(f"\n❌ 解压 {file_path.name} 失败：{e}")
        return False
    except Exception as e:
        print(f"\n❌ 解压 {file_path.name} 异常：{e}")
        return False


def extract_rar(file_path: Path, extract_dir: Path, password: str = None, progress_obj: ExtractProgress = None):
    """解压 rar 文件（适配中文路径+进度显示）"""
    try:
        rarfile.UNRAR_ENCODING = 'gbk'
        rarfile.UNRAR_TOOL = r"D:\My_App\normal_app_sys\winrar\UnRAR.exe"  # 请修改为你的UnRAR路径

        file_path_str = str(file_path.absolute())
        extract_dir_str = str((extract_dir / file_path.stem).absolute())
        progress_obj.current_file_name = file_path_str

        # 定义解压函数（用于进度跟踪）
        def extract_func(progress_hook):
            with rarfile.RarFile(file_path_str, 'r') as rf:
                if password:
                    rf.setpassword(password)
                os.makedirs(extract_dir_str, exist_ok=True)

                # 获取文件总数，按文件数估算进度
                file_list = rf.infolist()
                total_files = len(file_list)
                file_size = os.path.getsize(file_path_str)

                for i, file in enumerate(file_list):
                    try:
                        rf.extract(file, extract_dir_str)
                        # 按文件占比更新进度
                        unpacked = int((i + 1) / total_files * file_size)
                        progress_hook(unpacked)
                    except Exception as e:
                        print(f"\n⚠️  解压文件 {file.filename} 时出错：{e}")
                        continue

        # 跟踪解压进度
        track_extract_progress(file_path_str, extract_func, progress_obj)
        print(f"\n✅ 解压完成：{file_path.name} -> {extract_dir / file_path.stem}")
        return True
    except rarfile.BadRarFile:
        print(f"\n❌ 错误：{file_path.name} 不是有效的 RAR 文件")
        return False
    except rarfile.PasswordRequiredError:
        print(f"\n❌ 错误：{file_path.name} 需要解压密码")
        return False
    except rarfile.BadPassword:
        print(f"\n❌ 错误：{file_path.name} 解压密码错误")
        return False
    except FileNotFoundError:
        print(f"\n❌ 错误：未找到UnRAR.exe，请检查路径是否正确！当前配置路径：{rarfile.UNRAR_TOOL}")
        return False
    except Exception as e:
        print(f"\n❌ 解压 {file_path.name} 异常：{e}")
        return False


def extract_7z(file_path: Path, extract_dir: Path, password: str = None, progress_obj: ExtractProgress = None):
    """解压 7z 文件（适配中文路径+进度显示）"""
    try:
        file_path_str = str(file_path.absolute())
        target_dir = extract_dir / file_path.stem
        target_dir_str = str(target_dir.absolute())
        progress_obj.current_file_name = file_path_str

        # 定义解压函数（用于进度跟踪）
        def extract_func(progress_hook):
            kwargs = {}
            if password:
                kwargs['password'] = password

            with py7zr.SevenZipFile(file_path_str, mode='r', **kwargs) as zf:
                os.makedirs(target_dir_str, exist_ok=True)

                # 获取文件总数，按文件数估算进度
                file_list = zf.list()
                total_files = len(file_list)
                file_size = os.path.getsize(file_path_str)

                for i, file in enumerate(file_list):
                    try:
                        zf.extract(target_dir_str, [file.filename])
                        # 按文件占比更新进度
                        unpacked = int((i + 1) / total_files * file_size)
                        progress_hook(unpacked)
                    except Exception as e:
                        print(f"\n⚠️  解压文件 {file.filename} 时出错：{e}")
                        continue

        # 跟踪解压进度
        track_extract_progress(file_path_str, extract_func, progress_obj)
        print(f"\n✅ 解压完成：{file_path.name} -> {target_dir}")
        return True
    except py7zr.Bad7zFile:
        print(f"\n❌ 错误：{file_path.name} 不是有效的 7Z 文件")
        return False
    except py7zr.PasswordRequired:
        print(f"\n❌ 错误：{file_path.name} 需要解压密码")
        return False
    except py7zr.BadPassword:
        print(f"\n❌ 错误：{file_path.name} 解压密码错误")
        return False
    except Exception as e:
        print(f"\n❌ 解压 {file_path.name} 异常：{e}")
        return False


def batch_extract():
    """批量解压主函数（带完整进度显示）"""
    # 设置控制台编码
    os.system('chcp 65001 > nul')
    print("===== 批量解压压缩文件工具（适配中文路径+进度显示） =====")
    print("支持格式：.zip .rar .7z")
    print("===============================\n")

    # 1. 交互输入源目录
    while True:
        source_dir_input = input("请输入压缩文件所在目录路径：").strip()
        try:
            source_dir = validate_directory(source_dir_input)
            break
        except ValueError as e:
            print(f"❌ {e}，请重新输入！")

    # 2. 交互输入解压目标目录
    while True:
        extract_dir_input = input("请输入解压文件的目标目录路径：").strip()
        try:
            extract_dir = validate_directory(extract_dir_input)
            break
        except ValueError as e:
            print(f"❌ {e}，请重新输入！")

    # 3. 交互输入解压密码
    password = input("请输入统一解压密码（无密码直接回车）：").strip()

    # 4. 遍历压缩文件并计算总大小
    supported_formats = ('.zip', '.rar', '.7z')
    compressed_files = [
        f for f in source_dir.iterdir()
        if f.is_file() and f.suffix.lower() in supported_formats
    ]

    if not compressed_files:
        print("⚠️  指定目录下未找到 .zip/.rar/.7z 格式的压缩文件！")
        return

    # 计算总大小和剩余大小
    total_size = sum(os.path.getsize(f) for f in compressed_files)
    remaining_size = total_size

    # 初始化进度对象
    progress = ExtractProgress(
        total_files=len(compressed_files),
        current_index=0,
        remaining_files=len(compressed_files),
        total_size=total_size,
        current_file_size=0,
        current_file_unpacked=0,
        remaining_size=remaining_size
    )

    print(f"\n📌 共找到 {len(compressed_files)} 个压缩文件，总大小：{format_size(total_size)}")
    print("🚀 开始批量解压...\n")
    time.sleep(1)

    # 5. 逐个解压
    for idx, file in enumerate(compressed_files, 1):
        progress.current_index = idx
        progress.remaining_files = len(compressed_files) - idx
        progress.current_file_size = os.path.getsize(file)

        # 解压当前文件
        suffix = file.suffix.lower()
        success = False
        if suffix == '.zip':
            success = extract_zip(file, extract_dir, password, progress)
        elif suffix == '.rar':
            success = extract_rar(file, extract_dir, password, progress)
        elif suffix == '.7z':
            success = extract_7z(file, extract_dir, password, progress)

        # 更新剩余大小（仅当解压成功时）
        if success:
            remaining_size -= os.path.getsize(file)
            progress.remaining_size = remaining_size

    # 最终完成提示
    os.system('cls' if os.name == 'nt' else 'clear')
    print("🎉 批量解压任务执行完毕！")
    print(f"📊 总计处理：{len(compressed_files)} 个文件")
    print(f"📁 解压目录：{extract_dir}")
    print(f"📦 总大小：{format_size(total_size)}")


if __name__ == "__main__":
    batch_extract()
    input("\n按回车键退出...")  # 防止控制台闪退