@echo off
chcp 65001
"..\nx-spi-client\Asebis.Client.StarterCommand.exe" /u=nexus /p=fAvNCDnW3E /t=ImportLeistungen_CSV /o=100000000000000201 /f="2008-Le Pavec Anna+5860.csv" /map="..\..\HAS_map_main.csv" /v
Pause
