@REM a simple batch script which does some boring stuff

@echo off
goto :main

:copy_right
	echo.
	echo   === WW Pmod SMLD Control System ===
	echo   === powered by @ZL, 20210804    ===
	echo.
goto :eof

:main
	call :copy_right
	
	set path="D:\pj_00_codelib\2019_pypj\20210805 WW_Pmod_SMLD_Control_System\main.py"
	d:\DevEnv\WPy32-3741\python-3.7.4\python.exe %path%
	
	echo.
	pause
goto :eof