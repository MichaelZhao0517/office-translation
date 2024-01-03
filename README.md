# office-translation
Translate office files with Tencent API
Office文件自动翻译
介绍
基于腾讯云的Office文件翻译 支持新版的Office文件如xlsx,pptx,docx等 调用腾讯云的机器翻译API https://cloud.tencent.com/product/tmt 可以自行百度如何获取SecretId和SecretKey，填入即可使用

软件架构
Python（我是用3.7的版本做的，如果高于这个版本可以import包的有些方法会有所变化）

安装教程
安装Python
用pip安装必要的包（python-docx,腾讯云SDK，openpyxl,python-pptx)
运行.py文件即可，dist文件夹内有打包好的EXE文件可以直接使用
说明
本人非程序员，纯兴趣，目前Word当出现段落中嵌入图片的情况翻译后格式会有些错位，不过不影响使用， 如有大神可以帮忙完善一下最好~
