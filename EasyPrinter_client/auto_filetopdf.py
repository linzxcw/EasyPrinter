from win32com.client import DispatchEx
from win32com import client
import os
import time
import shutil
import yaml
# 转换doc,docx为pdf
def doc2pdf(fn):
      word = client.Dispatch("Word.Application")  # 打开word应用程序
      doc = word.Documents.Open(fn)  # 打开word文件

      a = os.path.split(fn)  # 分离路径和文件
      b = os.path.splitext(a[-1])[0]  # 拿到文件名

      doc.SaveAs("{}\\{}.pdf".format(path1, b), 17)  # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
      doc.Close()  # 关闭原来word文件
      word.Quit()



# 转换xls为pdf
def xls2pdf(fn):
      xlApp = DispatchEx("Excel.Application")
      xlApp.Visible = False
      xlApp.DisplayAlerts = 0
      books = xlApp.Workbooks.Open(fn,False)
      a = os.path.split(fn)  # 分离路径和文件
      b = os.path.splitext(a[-1])[0]  # 拿到文件名
      books.ExportAsFixedFormat(0, "{}\\{}.pdf".format(word_path, b))
      books.Close(False)
      xlApp.Quit()
      movefile("{}\\{}.pdf".format(word_path, b),path1)


# 获取指定路径下的所有word文件
# 可以穿透指定路径下的所有文件
def getfile(path):
    word_list = []  # 用来存储所有的word文件路径
    for current_folder, list_folders, files in os.walk(path):
        for f in files:  # 用来遍历所有的文件，只取文件名，不取路径名
            if f.endswith('doc') or f.endswith('docx') or f.endswith('xls') or f.endswith('xlsx') or f.endswith('pdf') or f.endswith('PDF'):  # 判断word文档
                word_list.append(current_folder + '\\' + f)  # 把路径添加到列表中
    return word_list  # 返回这个word文档的路径

#剪切文件从一个文件夹到另外一个文件夹
def movefile(fn,path1):
    filename = os.path.basename(fn)      #提取文件名
    tarpath = os.path.join(path1, filename)   #包含文件名的目标路径
    if os.path.exists(tarpath):
      os.remove(tarpath)
      shutil.move(fn, path1)
    else:
      shutil.move(fn, path1)

print("轻松打印客户端正在启动，请稍后......")
time.sleep(2)

with open('./config.yaml', 'r', encoding='utf8') as file:  # utf8可识别中文
    config = yaml.safe_load(file)
path1 = config['path0']  #服务端共享文件夹的路径
word_path =  os.path.abspath('.\\printclient')  # py文件所在目录下的打印文件夹，上传文件到此文件夹即转换成pdf  
print("客户端上传文件夹路径是：" + word_path)
print("服务端共享文件夹路径是：" + path1)
delete_pending = {}   


def docdocx2pdf():
    if __name__ == '__main__':
        print('[+] 转换中，请稍等……')
        #加载yaml文件中的路径
        words = getfile(word_path)
        for word in words:
            if word.endswith('doc'):
                doc2pdf(word)
            elif word.endswith('xls') or word.endswith('xlsx'):
                xls2pdf(word)
            elif word.endswith('pdf') or word.endswith('PDF'):
                 movefile(word,path1)
            else:
                doc2pdf(word)


print("运行中......")
print("*请把需要打印的文件发送到客户端的文件夹中，发送前请关闭office软件，避免出错")
while True:
    if len(os.listdir(word_path)) > 0:
        for file in os.listdir(word_path):
            if file not in delete_pending.keys():
                delete_pending[file] = time.time()
                try:
                    docdocx2pdf()
                except:
                    print("转换出错了")
                    print("继续转换...")
            else:
                if time.time() - delete_pending[file] > 20:
                    try:
                        os.remove(os.path.join(word_path, file))
                    except:
                        print("文件删除错误")
                    del delete_pending[file]
    time.sleep(1)
