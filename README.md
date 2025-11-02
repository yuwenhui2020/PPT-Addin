# PPT-Addin☛[最新版下载页](https://github.com/yuwenhui2020/PPT-Addin/releases)
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Fyuwenhui2020%2FPPT-Addin.svg?type=shield)](https://app.fossa.com/projects/git%2Bgithub.com%2Fyuwenhui2020%2FPPT-Addin?ref=badge_shield)

### 致开发者
编译环境：visual studio 2022
主要用到的东西：webviewer2
使用语言：C#
如果需要强签名密钥：12345678
强签名和原版（毕竟是我自己的证书和密钥）不一样，但仍可以自己修改编译（毕竟源码都给出来了）

### 关于开源
多功能版本也会开源，但是代码还在整理（肯定是免费啦）

### 这是什么？
这是一个用于PowerPoint和WPS的插件，可以在ppt内显示网页(自动打开)，也可以在局域网的网页中控制换页（需手动打开）
### 已支持的特性
ppt放映模式下的实时网页的显示

web控制台（可局域网翻页，普通点击和简易输入）

插入二维码（直接添加进当前ppt页面）

插入符号（支持聚焦到当前输入）

多显示屏正确显示（不会放错显示屏啦）

wps ppt和ms office powerpoint的支持

### 与同类产品比较
**任意修改**，本产品允许任何形式的修改和分发，只要别抹黑，咱们就是好朋友

**无需登录**，本产品不需要任何登录也不会获取任何有关用户的数据

**没有监控**，本产品不会像某同类产品一样获取并记录用户添加过的网页

**没有云控**，本产品不对用户进行云端控制，连自带控制都是局域网的

**完全免费**，本产品不需要您掏出一分钱即可直接使用

**无需授权**，本产品不需要征得作者的同意即可用于任何场景的使用

**功能正常**，经过十多个内部版本的迭代，目前已经没有会影响正常使用的bug了
### 使用前提条件
##### 系统要求
使用要求为win10以上的Office 2013及以上，安装了net 4.8和webview2

win7不保证可以使用，硬性重点是Office 2013及以上
##### webview2最新版（其实不一定是最新版）
点击[webview2离线安装包](https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/038e5be3-91a2-4c14-b2eb-2fac728c8c2c/MicrosoftEdgeWebView2RuntimeInstallerX86.exe)下载离线安装包
想在线安装则点击[webview2在线安装包](https://go.microsoft.com/fwlink/p/?LinkId=2124703)下载在线安装包。
##### .NET Framework 4.8及以上版本
点击[.NET Framework 4.8离线安装包](https://go.microsoft.com/fwlink/?linkid=2088631)下载
想在线安装则点击[.NET Framework 4.8在线安装包](https://go.microsoft.com/fwlink/?LinkId=2085155)下载在线安装包。
### 配置此插件
##### 安装
点击[下载页](https://github.com/yuwenhui2020/PPT-Addin/releases)选择最新的PowerPointAddIn.zip并下载，
然后在电脑的数据保存目录将压缩包解压，然后点击“安装工具.bat”进行快速安装
##### WPS安装
在[安装]的基础上，需要打开[WPS]-打开任意ppt文件-左上角[文件]-新菜单的[选项]-左侧[信任中心]-下方[受信任的加载项]-[启用所有第三方COM加载项，重启WPS后生效(E)]-勾选并确认，关闭WPS重启即可
##### 卸载
打开[PowerPoint]-[文件]菜单-右下角[选项]-左侧[加载项]-管理:COM加载项[转到]-左侧[PowerPointAddIn]-右侧[删除]
##### 更新
已安装过的只需要将新版的所有文件复制到之前目录覆盖即可，当然，重新[安装]到任意位置也可以
### 使用-插入网页
##### 插入在线网页
把想要显示的网页链接直接复制进插件的网址输入框，然后点击“添加网页”
##### 插入本地文件（绝对路径版）
使用浏览器打开本地的html文件或者pdf文件
此时地址栏显示的链接就可以粘贴到插件的网址输入框，然后点击“添加网页”
“此为相对路径”右侧的选择文件也可以添加准确的文件路径，当然这只是添加路径而已
##### 插入本地文件（相对路径版）
只需在上一条的基础上，点击“此为相对路径”即可，
然后点击“添加网页”，当然，记得把文件复制到与ppt同一目录

### 免责声明
本产品使用MIT开源协议，承诺**永远免费和无广告，以最精简的形式出现在大家眼中**
若有功能建议或是bug留言，请提issue，

### License
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Fyuwenhui2020%2FPPT-Addin.svg?type=large)](https://app.fossa.com/projects/git%2Bgithub.com%2Fyuwenhui2020%2FPPT-Addin?ref=badge_large)
