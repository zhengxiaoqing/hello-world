*** Settings ***
Library    AutoItLibrary       
#导入库
*** Test Cases ***

Test01
    AutoItLibrary.Run    calc.exe
    Wait For Active Window    计算器    
#显示计算器
Test02
     ${getcwd}    Evaluate    os.getcwd()    modules=os
