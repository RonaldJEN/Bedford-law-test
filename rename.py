import os
path = '/Users/PycharmProjects/Benfordslaw/000'#放置原始文件,更名前
for root, dirs, files in os.walk(path):
    for file in files:
        if file[0:5]=='现金流量表':
            file_path=os.path.join(root,file)
            if file_path.endswith('.xls'):
                new_file_name=file_path[-10:-4]+'cash.xls'
                os.rename(file_path,new_file_name)
            else:
                pass
        elif file[0:3]=='利润表':
            file_path=os.path.join(root,file)
            if file_path.endswith('.xls'):
                new_file_name=file_path[-10:-4]+'profit.xls'
                os.rename(file_path,new_file_name)
        else:
            file_path = os.path.join(root, file)
            if file_path.endswith('.xls'):
                new_file_name = file_path[-10:-4]+'asset.xls'
                os.rename(file_path, new_file_name)
