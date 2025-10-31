# auto_choose
## 一个自动做选择题的简易工具
目前大部分网站会对选择题内容进行前端加密，故采取`OCR`技术进行识别，本地部署`OCR`可以更便于不出网的情况。`api-key`换成自己的即可，这里也同样支持本地`ollama`调用。

## install
``` bash
# 下载并且于 Additional language data 中选择 Chinese 进行安装，安装结束后将安装的路径加入环境变量中
https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe 

# cmd运行
pip install -r requirements.txt

# 运行前修改 AI 模型配置
python run.py
```