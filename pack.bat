@echo off
setlocal

rem 定义文件和文件夹路径
set PACK_FOLDER=pack
set GUI_SCRIPT=gui.py

@REM rem 删除现有的pack文件夹及其内容
@REM if exist %PACK_FOLDER% (
@REM     rmdir /s /q %PACK_FOLDER%
@REM )

rem 新建pack文件夹
mkdir %PACK_FOLDER%

rem 执行打包操作
pyinstaller --onefile --noconsole %GUI_SCRIPT%

rem 移动生成的可执行文件到pack文件夹
move dist\gui.exe %PACK_FOLDER%

rem 删除临时生成的dist和build文件夹
rmdir /s /q dist
rmdir /s /q build

rem 提示操作完成
echo Packing completed.

endlocal
