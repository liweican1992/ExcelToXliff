#!/bin/sh

#  Localization.sh
#  UDictionary
#
#  Created by 李伟灿 on 2021/8/27.
#  Copyright © 2021 com.youdao. All rights reserved.


#英文文案录入 可以使用Excel里面的命令
#   =""""&A:A&""""&"="&""""&A:A&""""&";"
#chmod +x Localization.sh
#./Localization.sh
echo "=========开始执行========="


languages=(#"ar"
            "de"
            "en"
            "es"
            "fr"
            "id"
            #"it"
            "ja"
            "ko"
            "pt"
            #"ru"
            #"vi-VN"
            "zh-Hant"
            "zh-Hans"
            "th"
            )
outPath="$HOME/Desktop/exportLoaclization"
scheme="iTranscribe"


#导出xliff xcode10后为.xcloc格式文件夹
exportLocalizations(){
    echo "=========开始导出Xliff文件========="

    if test -e $outPath
    then
        echo "=========outPath existed, clean outPath"
        rm -rf $outPath
    fi

    exportLanguageStr=""

    for subLanguage in "${languages[@]}"
    do
        tmpStr="-exportLanguage $subLanguage "
        exportLanguageStr=$exportLanguageStr$tmpStr
    done
        #echo $exportLanguageStr

    #xcodebuild -exportLocalizations -project UDictionary.xcodeproj -localizationPath $HOME/Desktop/outData2 -exportLanguage en -exportLanguage ja

    #拼接命令
    commandStr="xcodebuild -exportLocalizations -project $scheme.xcodeproj -localizationPath $outPath $exportLanguageStr"

    echo $commandStr

    eval $commandStr
    echo "=========Xliff导出完成========="

}

runPython(){
    echo "=========执行Python脚本读取Excel========="
#    path=$HOME/Desktop/ExcelToXliff
#    cd $path
    chmod +x xliff.py
    python3 xliff.py
    echo "=========正在执行Python脚本========="

}

importLocalizations(){
    for sub in "${languages[@]}"
    do
        echo $sub

        tmpStr="-localizationPath $outPath/$sub.xcloc "
        commandStr="xcodebuild -importLocalizations -project $scheme.xcodeproj $tmpStr"
        echo $commandStr
        eval $commandStr
    done
        echo "=========Xliff导入完成========="
}


#导出xliff
exportLocalizations
#读取excel 写入xliff
runPython
#导入xliff
importLocalizations

