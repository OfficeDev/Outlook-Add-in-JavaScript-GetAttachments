---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: 此示例演示如何从 Exchange 邮箱中获取附件。
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:48:02 PM
---
# Outlook 加载项：从 Exchange Server 中获取附件

**目录**

* [摘要](#summary)
* [先决条件](#prerequisites)
* [示例主要组件](#components)
* [代码说明](#codedescription)
* [构建和调试](#build)
* [疑难解答](#troubleshooting)
* [问题和意见](#questions)
* [其他资源](#additional-resources)

<a name="summary"></a>
##摘要
此示例演示如何从 Exchange 邮箱中获取附件。

<a name="prerequisites"></a>
## 先决条件 ##

此示例要求如下：  

  - Visual Studio 2013 Update 5 或 Visual Studio 2015。  
  - 运行至少具有一个电子邮件帐户或 Office 365 帐户的 Exchange 2013 的计算机。你可以[参加 Office 365 开发人员计划并获取为期 1 年的免费 Office 365 订阅](https://aka.ms/devprogramsignup)。
  - 任何支持 ECMAScript 5.1、HTML5 和 CSS3 的浏览器，如 Internet Explorer 9、Chrome 13、Firefox 5、Safari 5.0.6 以及这些浏览器的更高版本。
  - 熟悉 JavaScript 编程和 Web 服务。

<a name="components"></a>
## 示例主要组件
本示例解决方案包含以下文件：

- AttachmentExampleManifest.xml：Outlook 加载项的清单文件。
- AppRead\Home\Home.html：Outlook 邮件加载项的 HTML 用户界面。
- AppRead\Home\Home.js：用于将附件信息发送到此示例中包含的远程附件服务的 JavaScript 文件。

AttachmentService 项目使用 WCF API 来定义 REST 服务。该项目包含以下文件：

- Controllers\AttachmentServiceController.cs：为示例服务提供业务逻辑的服务对象。
- Models\ServiceRequest：表示 Web 请求的对象。对象的内容通过从邮件加载项发送的 JSON 请求对象创建。
- Models\Attachment.cs：该实用工具对象有助于反序列化由邮件加载项发送的 JSON 对象。
- Models\AttachmentDetails.cs：表示每个附件的详细信息的对象。它提供了与邮件加载项的 `AttachmentDetails` 对象匹配的 .NET Framework 对象。
- Models\ServiceResponse：表示 Web 服务响应的对象。将对象的内容发送回邮件加载项时，会将其序列化为 JSON 对象。
- Web.config：将示例服务绑定到 Web 服务器终结点。



<a name="codedescription"></a>
##代码说明

此示例演示如何从支持邮件加载项的 Web 服务中检索附件。例如，你可以创建用于将照片上传到共享网站或将文档存储到存储库的服务。该服务可直接从 Exchange 服务器获取附件，无需客户端执行额外的操作来获取附件，然后将其发送到服务。

此示例分为两个部分。第一部分是在电子邮件客户端中运行的邮件应用。当邮件或约会是活动项目时，将显示邮件加载项。选择“**测试附件**”按钮时，邮件加载项会将有关附件的详细信息发送到处理请求的 Web 服务。该服务使用以下步骤来处理附件：

- 将 [GetAttachment](http://msdn.microsoft.com/library/aa494316(v=exchg.150).aspx) 操作请求发送到托管邮箱的 Exchange 服务器。服务器通过向此服务发送附件来响应。在此示例中，该服务仅从服务器写入 XML 以跟踪输出。
- 将已处理的附件数返回给邮件应用。



<a name="build"></a>
## 构建和调试 ##
**注意**：用户收件箱中具有一个或多个附件的任何电子邮件均会激活邮件加载项。在运行示例加载项之前，可以向测试帐户发送一封或多封电子邮件，以便更轻松地测试该加载项。

1. 打开 Visual Studio 中的解决方案。
2. 在解决方案资源管理器中，右键单击解决方案。选择“**设置启动项目**”。 
3. 选择“**通用属性**”，然后选择“**启动项目**”。
4. 确保将 **AttachmentExampleService** 项目的“**操作**”设置为“**启动**”。
5. 按 F5 生成并部署示例加载项。
6. 通过为 Exchange 2013 服务器提供电子邮件地址和密码连接至 Exchange 帐户。
7. 允许服务器配置邮件帐户。
8. 通过输入帐户名称和密码登录电子邮件帐户。 
9. 选择收件箱中的一封邮件。
10. 等待加载项栏出现在邮件上方。
11. 在加载项栏中，单击 **AttachmentExample**。
12. 当邮件加载项出现时，单击 **TestAttachments** 按钮以向 Exchange 服务器发送请求。
13. 服务器将响应为该项目处理的附件数。这应等于项目中包含的附件数。

<a name="troubleshooting"></a>
##疑难解答
以下是当你使用 Outlook Web App 测试 Outlook 的邮件加载项时可能发生的常见错误：

- 选中邮件后，不会显示加载项栏。如果发生此情况，请通过在 Visual Studio 窗口中选择“**调试 - 停止调试**”来重启应用程序，然后按 F5 重建并部署加载项。 
- 部署和运行加载项时，可能不会记录对 JavaScript 代码的更改。如果更改未记录，请清除 Web 浏览器上的缓存，方法是选择“**工具 - Internet 选项**”并单击“**删除…**”按钮。删除临时 Internet 文件，然后重启加载项。 

<a name="questions"></a>
##问题和意见##

- 如果你在运行此示例时遇到任何问题，请[记录问题](https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments/issues)。
- 与 Office 加载项开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)。确保你的问题或意见标记有 [Office 加载项]。


<a name="additional-resources"></a>
## 其他资源 ##

- [更多加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Web API：官方 Microsoft ASP.NET 网站](http://www.asp.net/web-api)
- [如何：从 Exchange 服务器获取附件](http://msdn.microsoft.com/library/dn148008.aspx)

## 版权信息
版权所有 (c) 2015 Microsoft。保留所有权利。


此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
