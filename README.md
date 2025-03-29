
# Outlook批量登录自动化工具

这是一个用于批量登录Outlook账号并自动访问Genspark邀请链接的工具。该工具提供了图形界面(GUI)和命令行两种使用方式，可以同时处理多个Outlook账号，提高工作效率。


# 将下图中的文件下载到一个文件夹里
![image](https://github.com/user-attachments/assets/4f45dc7c-3eca-493f-ba8a-4949555c12ea)

## 功能特点

- 批量登录多个Outlook账号
- 自动访问Genspark邀请链接
- 支持并发处理多个账号
- 提供图形界面和命令行两种使用方式
- 详细的日志记录
- CSV文件导入账号

## 安装步骤

### 1. 安装依赖

确保您的系统已安装Python 3.6或更高版本，然后安装所需的依赖包：

```bash
pip install -r requirements.txt
```

### 2. 安装Microsoft Edge浏览器

本工具使用Microsoft Edge浏览器进行自动化操作，请确保您的系统已安装最新版本的Edge浏览器。

### 3. 准备账号

账号格式为
- email: Outlook邮箱地址
- password: 对应的密码

示例格式：
```
email,password
example@outlook.com,yourpassword123
another@outlook.com,anotherpassword456
```

## 使用方法

### 图形界面(GUI)使用方法

1. 运行GUI程序：
# 直接在pycharm中run
或
```bash
python outlook_login_gui.py
```

2. GUI界面说明：

![GUI界面](https://example.com/gui_screenshot.png)

#### 界面各部分说明：

- **Genspark邀请链接**：输入您的Genspark邀请链接，默认已提供一个链接（需被覆盖掉）
- **最大并发数量**：设置同时处理的账号数量，建议根据电脑性能设置，建议为5
- **Outlook账号**：粘贴账号信息，格式为`email,password`，每行一个账号
- **从CSV导入**：点击此按钮从CSV文件导入账号
- **状态**：显示当前程序运行状态
- **运行日志**：显示程序运行的详细日志
- **运行**：开始执行自动化流程
- **停止**：停止当前运行的自动化流程
- **清除日志**：清除日志显示区域的内容
- **打开日志文件夹**：打开存储日志文件的文件夹

#### 操作步骤：

1. 输入或确认Genspark邀请链接
2. 设置最大并发数量
3. 直接在文本框中输入账号信息或通过「从CSV导入」按钮导入账号
4. 点击「运行」按钮开始自动化流程
5. 查看运行日志了解处理进度
6. 如需停止，点击「停止」按钮
7. 程序运行完成后，会弹出提示信息

### 命令行使用方法

如果您不需要图形界面，可以直接使用命令行运行自动化脚本：

```bash
python outlook_login_automation.py [最大并发数]
```

参数说明：
- `[最大并发数]`：可选参数，指定同时处理的账号数量，默认为5

#### 命令行使用示例：

```bash
# 使用默认并发数(5)运行
python outlook_login_automation.py

# 指定并发数为3运行
python outlook_login_automation.py 3
```

#### 命令行模式注意事项：

1. 确保当前目录下存在`outlook_accounts.csv`文件
2. 默认使用代码中设置的Genspark邀请链接
3. 如需修改邀请链接，请在代码中修改`GENSPARK_URL`变量
4. 日志将输出到控制台和`automation_log.txt`文件

## CSV文件格式要求

CSV文件必须包含以下列：
- `email`：Outlook邮箱地址
- `password`：对应的密码

文件第一行必须是列名，从第二行开始是账号数据。确保CSV文件使用UTF-8编码以支持特殊字符。

## 日志文件

- GUI模式下，日志保存在`logs`文件夹中，文件名格式为`outlook_automation_YYYYMMDD_HHMMSS.log`
- 命令行模式下，日志保存在`automation_log.txt`文件中

## 常见问题

### 1. 浏览器窗口没有自动关闭

这是正常现象。程序设计为保持浏览器窗口打开，以便您查看登录结果或进行手动操作。您可以在程序完成后手动关闭这些窗口。

### 2. 登录过程中出现验证码或安全验证

如果Microsoft要求额外的安全验证，程序会尝试处理常见的验证页面。对于需要手机验证码等操作，可能需要您手动干预。

### 3. 程序运行速度慢

尝试减少并发数量，过高的并发可能导致电脑性能下降或网络拥堵。

### 4. 导入CSV文件失败

确保CSV文件格式正确，包含必要的列名，并且使用UTF-8编码。

## 注意事项

- 本工具仅用于合法目的，请勿用于任何违反Microsoft服务条款的活动
- 大量自动登录可能触发Microsoft的安全机制，建议适度使用
- 请妥善保管您的账号信息，避免泄露

## 技术支持

如有问题或建议，请提交Issue或联系开发者。

## 许可证

[MIT License](LICENSE)
