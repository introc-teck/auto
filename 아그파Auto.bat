@echo off
set UserName=dpuser
set Passwd=gp14511
set Passwd2=dp7744

net use  * /delete
::net use m: /delete


net use x: \\172.16.33.253\OrderFile /user:%UserName% %Passwd%

net use w: \\172.16.33.253\Thum /user:%UserName% %Passwd%

net use y: \\172.16.33.216\HotFolderRoot\�Ʊ��� ó��_CMYK /user:%UserName% %Passwd2%

net use z: \\172.16.33.241\���� /user:%UserName% %Passwd%

