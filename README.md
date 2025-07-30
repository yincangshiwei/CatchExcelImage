
- - -

## 项目介绍

> 抓取Excel图片

## 功能介绍

### 软件截图

<img width="1001" height="1165" alt="image" src="https://github.com/user-attachments/assets/a802731d-b016-4a0a-96e3-15494b478ffb" />

### 基础功能

#### 文件选择

> 1. Excel文件：选择要处理的Excel文件
> 2. 输出目录：选择图片输出目录，默认用户桌面

#### 图片命名设置

> 说明：支持用户按组合规律命名和选择Excel列命名两种方式
> 1. 组合命名：
> 自定义：自定义固定输出名
> 包含日期：可设置不同日期格式进行输出，日期默认是当前时间
> 包含流水号：输出图片加上流水号，建议默认加上
> 顺序：可选择不同的组合顺序排列方式

> 2. Excel列命名：可选择用Excel某列或某几列做命名方式，一般是前缀命名，并且提供分隔符输入。

#### 提取模式

> 1. 整个工作簿：提取整个Excel文件里所有图片
> 2. 指定工作表： 填写sheet的名称，单独提取某个sheet的图片
> 3. 指定列：填写工作表名称和列（支持多个）来提取图片
> 4. 指定图片ID：一般只有嵌入图片方式才会用到，一般嵌入图片的内容是：=DISPIMG("ID_33BE28B532E140F9A441018DC3C8DB7D",1)。
> 5. 包含浮动图片：提取那些非嵌入的图片

## 相关技术

> Python版本：3.11.10(官方要求>=3.11)

> GUI UI使用：tkinter

> 打包程序使用：cx_Freeze

### 目录结构

```
BrowserUse/
├── gui.py
├── core.py
├── setup.py
├── requirements.txt
├── README.md
```

```

架构介绍：
- `gui.py`: gui界面代码。
- `core.py`: 主要逻辑代码。
- `setup.py`: 用于打包应用程序的脚本。
- `requirements.txt`: 列出了所有依赖项。
- `README.md`: 项目的说明文档（中文）。

```

### 二开说明

> 企业用户，可开启环境检测和修改内部网络地址，目前使用判断环境来控制非企业内部无法使用的处理，也可自行修改逻辑。

<img width="915" height="195" alt="image" src="https://github.com/user-attachments/assets/4e03029a-9beb-4514-b7df-901b601a9b43" />

<img width="885" height="686" alt="image" src="https://github.com/user-attachments/assets/6479d51d-899e-445f-8aa9-cee52b7bd6cf" />


## 使用说明

### 程序运行

> 运行文件：CatchExcelImage.exe

## 安装说明

### 1. 进入项目目录
```sh
cd CatchExcelImage
```

### 3. 安装相关依赖
```sh
pip install -r requirements.txt
```

### 4. 运行文件
```sh
python -u gui.py
```

### 5. 打包exe
```sh
python setup.py build

# 打包msi
# python setup.py bdist_msi
```
