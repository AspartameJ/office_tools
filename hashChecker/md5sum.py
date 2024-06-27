import hashlib
import json
import os

def calculate_md5(file_path):
    """计算文件的MD5哈希值"""
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def check_files_hash(json_file, root_directory):
    """检查目录及其子目录中的文件哈希值是否与JSON文件中的值匹配"""
    with open(json_file, 'r') as f:
        hash_dict = json.load(f)
    
    for root, _, files in os.walk(root_directory):
        for file_name in files:
            relative_path = os.path.relpath(os.path.join(root, file_name), root_directory)
            if relative_path in hash_dict:
                expected_md5 = hash_dict[relative_path]
                actual_md5 = calculate_md5(os.path.join(root_directory, relative_path))
                if actual_md5 == expected_md5:
                    print(f"{relative_path}: 哈希值匹配")
                else:
                    print(f"{relative_path}: 哈希值不匹配 (预期: {expected_md5}, 实际: {actual_md5})")
            else:
                print(f"{relative_path}: 未在哈希列表中找到")

if __name__ == "__main__":
    json_file = "hashes.json"  # 包含文件相对路径和MD5哈希值的JSON文件
    root_directory = "weights"  # 存放权重文件的根目录
    check_files_hash(json_file, root_directory)