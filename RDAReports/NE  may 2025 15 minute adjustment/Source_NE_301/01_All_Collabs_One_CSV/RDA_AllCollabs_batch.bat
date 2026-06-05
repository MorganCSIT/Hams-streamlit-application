@echo off
chcp 65001
"..\nx-spi-client\Asebis.Client.StarterCommand.exe" /u=nexus /p=fAvNCDnW3E /t=ImportLeistungen_CSV /o=100000000000000301 /f="RDA_AllCollabs+63936.csv" /map="..\HAS_map_main.csv" /v
Pause
