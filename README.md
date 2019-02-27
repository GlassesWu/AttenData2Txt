## 考勤文件转换脚本

 1. 因编码初期接收到的数据源文件为XLSX格式，所以主程序调用openpyxl库对Excel文件进行处理，后续使用中将XLS文件另存为XLSX格式即可
 2. 未使用pandas库进行快速处理，因为打包完的程序文件似乎更大（目前的打包完8M）
 3. 如有主程序变更需求请在完成后将主程序和图标重新打包，调用库为[Pyinstaller](https://pypi.org/project/PyInstaller/)
