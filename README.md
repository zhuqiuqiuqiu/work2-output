# Volvo实习产出
## 开发的自动配料-拣选app
![image](https://github.com/zhuqiuqiuqiu/work2-output/blob/main/app_page_pic.jpg)  

**说明**：    
1.下载exe程序到本地（桌面），当作app来使用  
2.选择零件-零件名-供应商数据：database，用来补充零件的Part_Name和供应商数据  
3.选择输入文件夹：这个文件夹里面放置从SAP导出来的各个料车的配料原始数据，参考案例  
4.选择输出文件夹：创建一个空文件夹，用于存放原始配料加工后的配料清单  

**project源码使用**：  
1.使用编辑器，Pycharm  
2.配置python解释器（conda）  

以下操作在**terminal终端**进行：  
3.下载代码打包成小程序的库  
pip install pyinstaller  
4.建立虚拟环境用于小程序打包，减少app的大小，让打包速度更快  
#创建名字为build_env 的虚拟环境  
python -m venv build_env  
#激活虚拟环境，在虚拟环境中进行操作  
build_env\Scripts\activate  
#安装必要的依赖  
pip install pandas openpyxl pyinstaller  
5.使用打包命令，在终端输入  
pyinstaller --onfile --windowed --name "配料清单批量生成工具" app_main.py  
可选命令：--icon:用于自定义  
