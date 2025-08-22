# VB6 自定义安装程序生成器 - 版本 2.0

我们很高兴地发布基于 VB6 的自定义安装程序生成器的重大更新——版本 2.0！此版本引入了强大的新功能，极大地提升了软件部署的灵活性和专业化水平。

## 🎉 新特性 (What's New)

### 1. 双重安装模式
现在支持两种安装模式以适应不同场景：
*   **正常模式（交互式）：** 提供完整的图形化安装向导界面，用户可浏览许可协议、选择安装目录并查看安装进度。
*   **静默模式（无交互）：** 支持通过命令行参数进行无人值守安装。非常适合系统管理员进行批量部署或集成到自动化脚本中。

### 2. 多语言支持
安装程序原生支持双语界面：
*   **英语 (English)**
*   **简体中文 (Simplified Chinese)**
安装界面将自动根据用户的操作系统语言设置显示对应的语言。

### 3. 动态配置文件 (`Setup.inf`)
所有安装设置现在均通过一个清晰的文本文件 `Setup.inf` 进行集中配置。您无需修改代码即可轻松定制安装包内容。

**可通过 `Setup.inf` 配置的内容包括：**
*   应用程序名称、版本号、开发商
*   默认安装路径
*   需要打包和安装的文件列表
*   是否创建开始菜单快捷方式和桌面图标
*   注册表项设置
*   ...以及更多选项。

## 🚀 使用方法 (How to Use)

1.  **配置:** 使用任何文本编辑器修改目录下的 `Setup.inf` 文件。
2.  **编译:** 在 VB6 环境中打开项目并编译生成您的定制化 `setup.exe`。
3.  **部署:** 分发生成的 `setup.exe` 和您的程序文件即可。

# VB6 Custom Installer Creator - Version 2.0

We are thrilled to announce a major update to our VB6-based Custom Installer Creator – Version 2.0! This release introduces powerful new features that significantly enhance the flexibility and professionalism of your software deployment process.

## 🎉 What's New

### 1. Dual Installation Modes
The installer now supports two distinct modes for different scenarios:
*   **Normal Mode (Interactive):** Provides a full graphical setup wizard interface. Users can browse the license agreement, choose the installation directory, and see the installation progress.
*   **Silent Mode (Unattended):** Supports command-line parameters for completely silent installation. Ideal for system administrators performing bulk deployments or integration into automated scripts.

### 2. Multi-Language Support
The installer natively supports dual-language interfaces:
*   **English**
*   **Simplified Chinese (简体中文)**
The installer interface will automatically display the corresponding language based on the user's operating system settings.

### 3. Dynamic Configuration File (`Setup.inf`)
All installation settings are now centrally configured through a clear text file: `Setup.inf`. You can easily customize your installation package without modifying the source code.

**What you can configure in `Setup.inf`:**
*   Application name, version, and vendor
*   Default installation path
*   List of files to be packaged and installed
*   Creation of Start Menu shortcuts and desktop icons
*   Registry settings
*   ...and many more options.

## 🚀 How to Use

1.  **Configure:** Edit the `Setup.inf` file in the directory using any text editor.
2.  **Compile:** Open the project in the VB6 environment and compile it to generate your customized `setup.exe`.
3.  **Deploy:** Distribute the generated `setup.exe` along with your application files.
