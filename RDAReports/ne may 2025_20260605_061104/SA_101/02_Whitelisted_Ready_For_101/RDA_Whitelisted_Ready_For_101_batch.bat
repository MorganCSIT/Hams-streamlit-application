@echo off
chcp 65001
"..\nx-spi-client\Asebis.Client.StarterCommand.exe" /u=nexus /p=fAvNCDnW3E /t=ImportLeistungen_CSV /o=100000000000000101 /f="RDA_Whitelisted_Ready_For_101+663.csv" /map="..\HAS_map.csv" /v
Pause
