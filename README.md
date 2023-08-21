# Wordgets: A word memory desktop gadget


### 


Wordgets是一个单词记忆桌面小工具。


==**ALERT**: Bug exists. ==


==警告：存在BUG。==


![wordgets Latest version](https://github.com/leaffeather/images/blob/main/wordgets-2.png?raw=true)
![image](https://github.com/leaffeather/wordgets/assets/134275609/891ef171-4a9b-41ee-b379-8e5d07166747)




This project was based on [**BeeWare**](https:github.com/beeware) - a cross-platform Python GUI library. It has the following features:


该项目之前基于BeeWare，这是一个跨平台的Python图形用户界面库。 此应用有如下特征：


*   **The rearend is supported**, which is the maximum motivation that I developed this app. The rearend part of rendering a template webpage is separated from the main application, which means that you can customize more functions in liberty. Here is the tutorial:


    支持后端。 这是我开发这个应用的最大动力。 渲染模板网页的后端部分从主程序分离出来，意味着你可以方便地个性化更多功能。 教程如下：


> 1.  Excel worksheet: You should backup the initial excel file and ensure it only have one worksheet. The row 1 will automatically be skipped, so you can put the corresponding column information in.
>
>     Excel表：你应当备份最初的Excel文件，并确保其中只有一张工作表。 第一行会自动忽略掉，因此，你可以写些有关列的信息。
> 2.  HTML template: Use `{{field}}` like cards in Anki. If you are unwilling to let the field content escape and embed text as webpage code, please use `{{field | safe}}` in the webpage.
>
>     网页模板：像Anki卡片一样使用{{field}}（field：字段）。 如果你不希望field的内容转义并且将其作为网页代码的一部分，请在网页中使用{{field | safe}}。
> 3.  Python script\:Please follow the example:
>
>     Python脚本：参考如下：
> 4.  If you wish that the word pronounciates automatically, you should use like `<audio src="..." **autopplay**></audio>`
>     
>     如果你想要单词自动发音，你应当使用例如`<audio src="..." **autopplay**></audio>`


```python
# Import
exec("from jinja2 import Environment, FileSystemLoader", globals())  # Immutable
exec("...")                 # Place the import statement at '...'


# Create variables
field = A           # field for html template and A is the column title in the Excel
field_2 = B         # An example for another field
...                 # More fields


# Immutable block
searchpath = folder_path
name = file_name
env = Environment(loader=FileSystemLoader(searchpath=searchpath))
template = env.get_template(name=name)


# You can customize the backend here
...


# It must place at the end
output = template.render(variable_name = variable_name,
                         variable_name_2 = variable_name_2,
                         ... )  # Copy the used fields twice, add a "=" sign in the middle
```


> **If you use it on a moblie OS, please use http/https to import data.**
> 
> **如果你使用的是移动操作系统来运行它，请使用http/https的方式导入数据**


*   Support **customizing one or two types of cards with different number of sides** via a simple Management.


    支持自定义面数不同的一到两种卡。
*   Change between different word lists in just a ComboBox and automatically load the vocabulary card when the app is opened.


    在下拉框里切换不同的单词卡，并在程序打开时自动加载单词卡。
*   Synchonize learning situation via Baidu NetDisk


    利用百度云同步学习情况



Known problems:


已知问题：


*   **Be lag when changing cards**. The bigger the size of the used Excel file is, the slower the next cards displays. However, when the Excel saves as .csv by Excel 2021 or WPS 2023 I tested, the phonetic symbols become messy codes even if encoded again and that is why I does not use this way though .csv is read much faster than .xlsx in Python. ** Recommend multiple split wordlists instead of the whole one**# Wordgets: A word memory desktop gadget


### 


Wordgets是一个单词记忆桌面小工具。


==**ALERT**: Bug exists. ==


==警告：存在BUG。==


![wordgets Latest version](https://github.com/leaffeather/images/blob/main/wordgets-2.png?raw=true)
![image](https://github.com/leaffeather/wordgets/assets/134275609/891ef171-4a9b-41ee-b379-8e5d07166747)




This project was based on [**BeeWare**](https:github.com/beeware) - a cross-platform Python GUI library. It has the following features:


该项目之前基于BeeWare，这是一个跨平台的Python图形用户界面库。 此应用有如下特征：


*   **The rearend is supported**, which is the maximum motivation that I developed this app. The rearend part of rendering a template webpage is separated from the main application, which means that you can customize more functions in liberty. Here is the tutorial:


    支持后端。 这是我开发这个应用的最大动力。 渲染模板网页的后端部分从主程序分离出来，意味着你可以方便地个性化更多功能。 教程如下：


> 1.  Excel worksheet: You should backup the initial excel file and ensure it only have one worksheet. The row 1 will automatically be skipped, so you can put the corresponding column information in.
>
>     Excel表：你应当备份最初的Excel文件，并确保其中只有一张工作表。 第一行会自动忽略掉，因此，你可以写些有关列的信息。
> 2.  HTML template: Use `{{field}}` like cards in Anki. If you are unwilling to let the field content escape and embed text as webpage code, please use `{{field | safe}}` in the webpage.
>
>     网页模板：像Anki卡片一样使用{{field}}（field：字段）。 如果你不希望field的内容转义并且将其作为网页代码的一部分，请在网页中使用{{field | safe}}。
> 3.  Python script\:Please follow the example:
>
>     Python脚本：参考如下：
> 4.  If you wish that the word pronounciates automatically, you should use like `<audio src="..." **autopplay**></audio>`
>
>     如果你想要单词自动发音，你应当使用如`<audio src="..." **autopplay**></audio>`


```python
# Import
exec("from jinja2 import Environment, FileSystemLoader", globals())  # Immutable
exec("...")                 # Place the import statement at '...'


# Create variables
field = A           # field for html template and A is the column title in the Excel
field_2 = B         # An example for another field
...                 # More fields


# Immutable block
searchpath = folder_path
name = file_name
env = Environment(loader=FileSystemLoader(searchpath=searchpath))
template = env.get_template(name=name)


# You can customize the backend here
...


# It must place at the end
output = template.render(variable_name = variable_name,
                         variable_name_2 = variable_name_2,
                         ... )  # Copy the used fields twice, add a "=" sign in the middle
```


> **If you use it on a moblie OS, please use http/https to import data.**
> 
> **如果你使用的是移动操作系统来运行它，请使用http/https的方式导入数据**


*   Support **customizing one or two types of cards with different number of sides** via a simple Management.


    支持自定义面数不同的一到两种卡。
*   Change between different word lists in just a ComboBox and automatically load the vocabulary card when the app is opened.


    在下拉框里切换不同的单词卡，并在程序打开时自动加载单词卡。
*   Synchonize learning situation via Baidu NetDisk


    利用百度云同步学习情况



Known problems:


已知问题：


*   **Be lag when changing cards**. The bigger the size of the used Excel file is, the slower the next cards displays. However, when the Excel saves as .csv by Excel 2021 or WPS 2023 I tested, the phonetic symbols become messy codes even if encoded again and that is why I does not use this way though .csv is read much faster than .xlsx in Python. ** Recommend multiple split wordlists instead of the whole one**

```python
# Import
exec("from jinja2 import Environment, FileSystemLoader", globals())  # Immutable
exec("...")                 # Place the import statement at '...'

# Create variables
field = A           # field for html template and A is the column title in the Excel
field_2 = B         # An example for another field
...                 # More fields

# Immutable block
searchpath = folder_path
name = file_name
env = Environment(loader=FileSystemLoader(searchpath=searchpath))
template = env.get_template(name=name)

# You can customize the backend here
...

# It must place at the end
output = template.render(variable_name = variable_name,
                         variable_name_2 = variable_name_2,
                         ... )  # Copy the used fields twice, add a "=" sign in the middle
```

> **If you use it on a moblie OS, please use http/https to import data.**
> 
> **如果你使用的是移动操作系统来运行它，请使用http/https的方式导入数据**

*   Support **customizing one or two types of cards with different number of sides** via a simple Management.

    支持自定义面数不同的一到两种卡。
*   Change between different word lists in just a ComboBox and automatically load the vocabulary card when the app is opened.

    在下拉框里切换不同的单词卡，并在程序打开时自动加载单词卡。
*   Synchonize learning situation via Baidu NetDisk

    利用百度云同步学习情况
    

Known problems:

已知问题：

*   **Be lag when changing cards**. The bigger the size of the used Excel file is, the slower the next cards displays. However, when the Excel saves as .csv by Excel 2021 or WPS 2023 I tested, the phonetic symbols become messy codes even if encoded again and that is why I does not use this way though .csv is read much faster than .xlsx in Python. ** Recommend multiple split wordlists instead of the whole one**

    切换单词卡有卡顿。使用的Excel文件越大，显示下一张卡片也就越慢。然而，我测试的在Excel 2021或WPS2023里将Excel另存为.csv，音标会成乱码，即使重新编码也是一样。因此，尽管python读取.csv文件远远快于.xlsx，我也没采用这种方式。**建议将一整个词表分成几个**
    
*   Scroll bars in Management does not work in Windows. Currently, when the information of a card or a word list you add just displays abnormally, you should change the window size of Management.

    管理器的滚动条在Windows上失效。当前，当你添加的卡片或者单词表信息显示异常，你需要调整管理器的窗口大小。
*   There could be something unreasonable in the used memory algorithm [SuperMemo2](https://pypi.org/project/supermemo2/). The selection of feeling strange(quality=0) and vague(quality=2) will return the same next review date (i.e. 1 day later). Currently the solution is that a word with quality=0 should be review today. Also, I have not seen what regulation should be adopted to compute the review priority of the studied words. Currently, I use a policy that GPT recommends: the former the review date, or the smaller the value of the expression $` easiness^{repetitions}/{interval} `$ with the same review date, the urgent the review of the word.

    使用的SuperMemo2算法可能有些不合理的地方。选择感到陌生（quality=0）和模糊（quality=2）得到的下一次复习日期是一致的（即都是一天后）。当前的解决方法是quality=0时，今天之内应当再次复习。此外，我没有看到如何计算学习过的单词的优先级的规则。当前，我使用的是GPT推荐的策略：复习日期越早或者相同日期的情况下表达式$` easiness^{repetitions}/{interval} `$的值越小，词的复习越紧急。


Future plans (may not realize for heavy work):

未来计划（工作繁重，也许并不会实现）：

*   To add a launcher for Desktop OSs to hide the title bar and make the main program always on top, which can realize the style just like Windows Vista gadgets.

    为桌面操作系统平台加一个启动器，用于隐藏标题栏并让程序始终置顶，实现类似于Windows Vista系统的小工具的样式。

Welcome developers who want to help improve! Contact [leaffeather@foxmail.com](mailto://leaffeather@foxmail.com)

欢迎任何想帮忙改进的开发者！请联系[leaffeather@foxmail.com](mailto://leaffeather@foxmail.com)
