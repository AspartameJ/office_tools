import os

def generate_image_paths_file(base_folder, subsets, output_folder):
    for subset in subsets:
        coco_subset_folder = os.path.join(base_folder, subset)
        output_file = os.path.join(output_folder, f'{subset}2017.txt')

        # 打开文件准备写入图片路径
        with open(output_file, 'w') as file:
            # 遍历文件夹中的所有图片文件
            for filename in os.listdir(coco_subset_folder):
                if filename.endswith(('.jpg', '.jpeg', '.png')):
                    file_path = os.path.join(coco_subset_folder, filename)
                    file.write(file_path + '\n')

        print(f'Image paths for {subset} saved to {output_file}')

# 示例使用
base_folder = "/home/nfs/appnfs/alipay_00000559/datacopy/coco/images"
subsets = ['train', 'val']
output_folder = "/home/nfs/appnfs/alipay_00000559/datacopy/coco"
generate_image_paths_file(base_folder, subsets, output_folder)