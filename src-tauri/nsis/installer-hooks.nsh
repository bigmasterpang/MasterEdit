; MasterEdit 安装钩子
; 安装完成后写入标记文件，应用中用于识别「安装版」（安装版走安装包静默升级，绿色版走自替换）
; 另外尽力清理改名前的旧可执行文件（mastermd.exe / MasterMD.exe），避免安装目录里同时留下新旧两个程序
; 注意：Delete 失败不会中断安装；若旧程序仍在运行导致文件被占用，会由下一次安装或用户手动删除

!macro NSIS_HOOK_POSTINSTALL
  FileOpen $0 "$INSTDIR\installed.marker" w
  FileWrite $0 "installed"
  FileClose $0
  Delete "$INSTDIR\mastermd.exe"
  Delete "$INSTDIR\MasterMD.exe"
!macroend
