:: C:\>dir MSVCR120_CLR0400.dll /s /b
:: C:\Windows\System32\msvcr120_clr0400.dll
:: C:\Windows\SysWOW64\msvcr120_clr0400.dll

set DOWHAT=Reinstate
c:
cd\Windows
if %DOWHAT%==Backup (
    ren c:\Windows\System32\msvcr120_clr0400.dll msvcr120_clr0400.dll.bak
    ren c:\Windows\SysWOW64\msvcr120_clr0400.dll msvcr120_clr0400.dll.bak
) else (
    ren c:\Windows\System32\msvcr120_clr0400.dll.bak \msvcr120_clr0400.dll
    ren c:\Windows\SysWOW64\msvcr120_clr0400.dll.bak msvcr120_clr0400.dll
)

pause


