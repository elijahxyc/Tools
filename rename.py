import os

path = 'D:\\Downloads'

def file_name(file_dir):
    L=[]   
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if os.path.splitext(file)[1] == '.flac':  
                L.append(os.path.join(root, file))
    return L

def rename(file_dir):
    i=0
    L=[]
    filelist=os.listdir(file_dir)#该文件夹下所有的文件（包括文件夹）
    for files in filelist:#遍历所有文件
        i=i+1
        Olddir=os.path.join(path,files) #原来的文件路径                
        if os.path.isdir(Olddir):#如果是文件夹则跳过
            continue
        filename=os.path.splitext(files)[0] #文件名
        filetype=os.path.splitext(files)[1] #文件扩展名
        Newdir=os.path.join(file_dir,filename+filetype) #新的文件路径
        if filename.find('风Q长林') != -1:
            filename_temp = filename.replace('1080p高清国语中字','')
            #print(filename_temp)
            for i in range(1,51):
                strEp = ''
                if i >= 1 and i <= 9:
                    strEp = '0' + str(i)
                else:
                    strEp = str(i)
                
                if filename_temp.find(strEp) != -1:
                    filename_new = 'fqcl' + strEp
                    Newdir = os.path.join(file_dir,filename_new+filetype)
                    break
                
        #Newdir = Newdir.replace(' ', '') 
        print(Newdir)
        L.append(Newdir)
        #os.rename(Olddir,Newdir)#重命名
    
    return L

rename(path)
