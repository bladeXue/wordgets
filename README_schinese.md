# Wordgets: 一个桌面背单词软件



### 

![cover](https://github.com/leaffeather/images/blob/main/wordgets_cover.png?raw=true)


**更新日志：**

v1.0.1:

(1)修复了一些错误的消息框。（在macOS上，如果你使用了网络代理或者网络未连接，问题对话框不允许你做出任何选择。请在解决后重启软件。）

---
1.  简介

    Wordgets是一个桌面背单词软件，基于跨平台Python原生框架的[Beeware](https://beeware.org/)开发。具有如下特点：
    - **支持自定义后端**。注意：使用者通过后端使用网络服务时<u>**必须遵守当地网络法律法规和有关的网络服务提供者/商的条例**</u>。
    - **支持通过百度网盘实现跨设备、跨平台的同步**。 
    - **支持自动播放音频、视频**。
    
    受支持的操作系统:
    - Windows >= 10
    - Android >= 10
    - MacOS >= 10.10
2.  使用说明
    1. 关于单词表
    
        单词表应当为一个只有一张工作表的.xlsx Excel工作簿文件。第一行放列名，随后一行行放单词和有关的内容。建议每个单词表在最多放大约4000个词，否则程序的使用过程中会有明显卡顿。
    2. 关于模板卡
    
        这一部分需要少量的编程知识。
        你可以为单词表指定一到两种单面或双面的模板卡。 每张卡面包含前端.html模板网页和后端.py Python脚本。具体来说：
        1. 模板网页（所有支持的操作系统）
       
            创建一个HTML，使用字段将其包裹在两对花括号之间，如字段*word*写成`{{word}}`，字段*explanation*写成`{{explanation}}`。如果你想嵌入用python生成的HTML代码 ，你需要使用字段以字段*formula*写作`{{formula | safe}}`的方式。启用自动朗读则应使用例如`<audio src="[音频地址]" autoplay></audio>`的方式。
        
        2. python脚本（Windows/Android上）
            ```python
            exec("from jinja2 import Environment, FileSystemLoader",globals())  # ** 不要改 **
            exec("from bs4 import BeautifulSoup",globals())   # 用这种方式导入你想用的包
            
            # 左边是字段名，右边是Excel中字母形式的列序号
            word = A 
            explanation = B 
            
            searchpath = folder_path                                            # ** 不要改 ** 
            name = file_name                                                    # ** 不要改 ** 
            env = Environment(loader=FileSystemLoader(searchpath=searchpath))   # ** 不要改 ** 
            template = env.get_template(name=name)                              # ** 不要改 ** 
            
            # 你可以自定义一些代码在这个地方
            formula = "<div id="raw-formula"> \(\int_a^b f(x)dx\) </div>"
            
            # 你只需要修改括号里的内容，
            # 将字段复制两次，中间添个=，记得补上逗号
            output = template.render(word = word, 
                                     explanation = explanation,
                                     formula = formula)
            ```
        3. python脚本（MacOS上）
            ```python
            exec("from flask import Flask, render_template", globals())         # ** 不要改 **
            exec("from bs4 import BeautifulSoup",globals())   # 用这种方式导入你想用的包
            
            # 左边是字段名，右边是Excel中字母形式的列序号
            word = A 
            explanation = B 
            
            searchpath = folder_path                                            # ** 不要改 **
            name = file_name                                                    # ** 不要改 **
            app = Flask(__name__, template_folder = searchpath)                 # ** 不要改 **
            
            # 你可以自定义一些代码在这个地方
            formula = "<div id="raw-formula"> \(\int_a^b f(x)dx\) </div>"
            
            @app.route('/')                                                     # ** 不要改 **
            def index():                                                        # ** 不要改 **
                global word, explanation, formula # 在global后面放上你用到的全部字段，
                                                  # 每两个之间添个逗号
                
                # 不要改第一行
                # 把字段复制两次，添个=在中间，
                # 别忘记添加逗号
                return render_template(template_name_or_list = name, 
                                       word = word, 
                                       explanation = explanation,
                                       formula = formula)
                                
            app.run(debug = False, use_reloader = False)                        # ** 不要动 **
            ```
    3. 如何导入单词表与单词卡
    
        如果你使用的是Windows或MacOS，你可以简单地在输入框内填上目标文件的绝对路径。此应用总是会和这些文件交互，因而不要动它们。
        对于Android，事情变得有些复杂。请按照如下步骤：
            1. 你的手机和Windows电脑应处在同一局域网内。
            2. 下载已经配置过了的nginx并解压Download the already configured nginx and unzip it. 
            3. Place your wordlists and cards into directory `WWW\` of nginx. 运行nginx.exe，它不会有窗口产生，因此只要确保你运行了就行。
            4. 打开`cmd`，输入`ipconfig`回车。找到你的电脑局域网ip地址。
            5. 导入文件用`http://[你的IP地址]:8080/[文件名].[后缀名]`
    4. 如何使用同步
    
        请注册一个百度账号并且至少登陆过百度网盘一次。然后，按照下面的步骤： 
        1. 打开[百度网盘开放平台](https://pan.baidu.com/union); 
        2. 点击右上角[申请加入]，完成申请; 
        3. 点击[控制台]，创建一个应用; 
        4. 查看应用详情，复制[AppKey]填到wordgets的登录步骤中相应的输入框中。
        
        注意:
        1. 你最好在别的设备上同步完成后再打开软件。
        2. 当你不想继续背单词时，你用当关闭软件。
        3. 如果同步失败并且消息提示你无网络连接，有可能是代理应用例如Clash正在运行的缘故。
        4. Wordgets会在你的百度网盘中创建`wordgets_sync`目录，不要动！除非你打算重置同步数据库。

3.  已知问题
    1.  [windows上] 设置内的滚动条失灵，你可以通过拉伸窗口来显示完全。
    2.  [MacOS上] 窗口组件自动调整失灵。
