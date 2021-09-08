# ExcelToXliff

## 背景
    
目前负责的海外项目支持十几种语言的国际化

Apple使用Xliff管理国际化文案，虽然可以直接交付Xliff文件给人工翻译组。

但是考虑到安卓端以及方便管理，目前翻译结果依旧是Excel格式。

而Xliff为XML格式。所以需要进行文案录入。之前方案为人工使用软件XliffTool进行录入

缺点是容易出错以及工作量大。
    
所以用Python3.8写了一个批量录入脚本

update:新增了shell脚本

支持批量从xcode导出xliff
执行python脚本批量录入文案
批量导入Xcode

## 使用

1、确保excel在桌面，文件名为localizable.xls

2、文案表名为ios或者为第一个表

3、执行Localization.sh脚本

执行方法

打开terminal在开发目录或者Localization.sh目录下

```
chmod +x Localization.sh
./Localization.sh
```



运行Python需要xlrd库

![image](1.png)

Excel格式如图

![image](2.png)

