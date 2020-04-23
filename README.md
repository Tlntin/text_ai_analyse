## 代码使用说明

1. 安装依赖
```python
pip install -r requirement.txt
```
2. 打包成exe
```python
pyinstaller --paths=Deja_Vu_Sans_Mono.ttf -F -w -i my.ico main.py
```

## 主要功能
- 提取word中的文字，并且按句号进行分句
- 然后调用百度的AI对每一句话进行分析，找错可能错误的地方
- 文件-保存配置，可以把填入的api信息直接保存到本地
- 文件-加载配置，可以把本地的配置直接填入GUI
## 详细使用说明
- 请看软件， 菜单-使用说明