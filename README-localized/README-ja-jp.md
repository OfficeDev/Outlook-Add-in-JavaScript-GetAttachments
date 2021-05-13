---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: このサンプルでは、Exchange メールボックスから添付ファイルを取得する方法を説明します。
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:48:02 PM
---
# Outlook アドイン:Exchange Server から添付ファイルを取得する

**目次**

* [概要](#summary)
* [前提条件](#prerequisites)
* [サンプルの主要なコンポーネント](#components)
* [コードの説明](#codedescription)
* [ビルドとデバッグ](#build)
* [トラブルシューティング](#troubleshooting)
* [質問とコメント](#questions)
* [その他のリソース](#additional-resources)

<a name="summary"></a>
##概要
このサンプルでは、Exchange メールボックスから添付ファイルを取得する方法を説明します。

<a name="prerequisites"></a>
## 前提条件 ##

このサンプルを実行するには次のものが必要です。  

  - Visual Studio 2013 更新プログラム 5 または Visual Studio 2015。  
  - 少なくとも 1 つのメール アカウントまたは Office 365 アカウントがある Exchange 2013 を実行するコンピューター。[Office 365 Developer プログラムに参加すると、Office 365 の 1 年間無料のサブスクリプションを取得](https://aka.ms/devprogramsignup)できます。
  - Internet Explorer 9、Chrome 13、Firefox 5、Safari 5.0.6、またはこれらのブラウザーの最新バージョンなど、ECMAScript 5.1、HTML5、CSS3 をサポートするブラウザー。
  - JavaScript プログラミングと Web サービスに関する知識。

<a name="components"></a>
## サンプルの主な構成要素
このサンプル ソリューションに含まれるファイルは次のとおりです。

- AttachmentExampleManifest.xml:Outlook アドインのマニフェスト ファイル。
- AppRead\Home\Home.html:Outlook 用メール アドインの HTML ユーザー インターフェイス。
- AppRead\Home\Home.js:添付ファイル情報を、このサンプルに含まれているリモート添付サービスに送信する JavaScript ファイル。

AttachmentService プロジェクトは、WCF API を使用して REST サービスを定義します。プロジェクトには次のファイルが含まれます。

- Controllers\AttachmentServiceController.cs:サンプル サービスのビジネス ロジックを提供するサービス オブジェクト。
- Models\ServiceRequest:Web 要求を表すオブジェクト。オブジェクトの内容は、メール アドインから送信された JSON 要求オブジェクトから作成されます。
- Models\Attachment.cs:メール アドインによって送信される JSON オブジェクトの逆シリアル化をできるようにするユーティリティ オブジェクト。
- Models\AttachmentDetails.cs:添付ファイルの詳細を表すオブジェクト。メール アドインの `AttachmentDetails` オブジェクトに一致する .NET Framework オブジェクトが提供されます。
- Models\ServiceResponse:Web サービスからの応答を表すオブジェクト。オブジェクトの内容は、メール アドインに送り返された際に、JSON オブジェクトにシリアル化されます。
- Web.config:サンプル サービスを Web サーバー エンドポイントにバインドします。



<a name="codedescription"></a>
##コードの説明

このサンプルでは、メール アドインをサポートする Web サービスから添付ファイルを取得する方法を説明します。たとえば、共有サイトに写真をアップロードするサービスや、ドキュメントをリポジトリに保存するサービスを作成できます。このサービスは、Exchange Server から直接添付ファイルを取得します。添付ファイルを取得してサービスに送信するためにクライアントが処理を実行する必要はありません。

このサンプルは 2 つのパーツから成ります。最初のパーツはメール アプリで、メール クライアントで実行されます。メッセージや予定がアクティブなアイテムの場合は、メール アドインが表示されます。[**添付ファイルのテスト**] ボタンを選択すると、メール アドインは、要求を処理する Web サービスに添付ファイルに関する詳細を送信します。サービスは、次の手順で添付ファイルを処理します。

- [GetAttachment](http://msdn.microsoft.com/library/aa494316(v=exchg.150).aspx) 操作の要求をメールボックスをホストする Exchange Server へ送信します。サーバーは、サービスへ添付ファイルを送信して応答します。このサンプルでは、サービスはサーバーから XML を書き込み、出力を追跡します。
- メール アプリに対して処理された添付ファイルの数を返します。



<a name="build"></a>
## ビルドとデバッグ ##
**注**: メール アドインは、1 つ以上の添付ファイルがある受信トレイのすべてのメール メッセージで有効になります。サンプル アドインを実行する前に、1 つ以上のメール メッセージをテスト アカウントに送信しておくと、より簡単にアドインをテストできます。

1. Visual Studio でソリューションを開きます。
2. ソリューション エクスプローラーで、ソリューションを右クリックします。[**スタートアップ プロジェクトの設定**] を選択します。 
3. [**共通プロパティ**] を選択して、[**スタートアップ プロジェクト**] を選択します。
4. **AttachmentExampleService** プロジェクトの **アクション** が **開始** に設定されていることを確認します。
5. F5 キーを押して、サンプル アドインをビルドおよび展開します。
6. Exchange 2013 Server 用のメール アドレスとパスワードを入力して Exchange アカウントに接続します。
7. サーバーがメール アカウントを構成できるようにします。
8. アカウント名とパスワードを入力して、メール アカウントにログオンします。 
9. 受信トレイのメッセージを選択します。
10. メッセージにアドイン バーが表示されるまで待ちます。
11. アドイン バーの [**AttachmentExample**] をクリックします。
12. メール アドインが表示されたら、[**TestAttachments**] ボタンをクリックして、Exchange Server に要求を送信します。
13. サーバーは、アイテムに対して処理された添付ファイルの数で応答します。これは、アイテムに含まれる添付ファイルの数と同じである必要があります。

<a name="troubleshooting"></a>
##トラブルシューティング
Outlook Web App を使用して Outlook のメール アドインをテストするときに発生する可能性がある一般的なエラーは次のとおりです。

- メッセージが選択されているときに、アドイン バーが表示されない。この問題が発生した場合は、Visual Studio ウィンドウで **[デバッグ]、[デバッグの停止]** の順に選択してアプリケーションを再起動し、次に F5 キーを押してアドインをリビルドして展開します。 
- アドインの展開と実行時に JavaScript コードの変更が認識されない場合がある。変更が認識されない場合は、**[ツール]、[インターネット オプション]** の順に選択し、[**削除**] ボタンを選択して Web ブラウザーのキャッシュをクリアします。インターネット一時ファイルを削除してからアドインを再起動します。 

<a name="questions"></a>
##質問とコメント##

- このサンプルの実行について問題がある場合は、[問題をログに記録](https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments/issues)してください。
- Office アドイン開発全般の質問については、「[Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)」に投稿してください。質問やコメントには、必ず "office-addins" のタグを付けてください。


<a name="additional-resources"></a>
## その他のリソース ##

- [その他のアドイン サンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Web API:Microsoft ASP.NET の公式サイト](http://www.asp.net/web-api)
- [操作方法:Exchange Server から添付ファイルを取得する](http://msdn.microsoft.com/library/dn148008.aspx)

## 著作権
Copyright (c) 2015 Microsoft.All rights reserved.


このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
