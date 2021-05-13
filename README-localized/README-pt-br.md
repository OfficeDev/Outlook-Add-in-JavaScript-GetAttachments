---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: Esse exemplo mostra como obter anexos de uma caixa de correio do Exchange.
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:48:02 PM
---
# Suplemento do Outlook: Obter os anexos de um servidor Exchange

**Sumário**

* [Resumo](#summary)
* [Pré-requisitos](#prerequisites)
* [Componentes principais do exemplo](#components)
* [Descrição do código](#codedescription)
* [Criar e depurar](#build)
* [Solução de problemas](#troubleshooting)
* [Perguntas e comentários](#questions)
* [Recursos adicionais](#additional-resources)

<a name="summary"></a>
##Summary
Este exemplo mostra como obter anexos de uma caixa de correio do Exchange.

<a name="prerequisites"></a>
## Pré-requisitos ##

Esse exemplo requer o seguinte:  

  - Visual Studio 2013 com Atualização 5 ou Visual Studio 2015.  
  - Um computador executa o Exchange 2013 com pelo menos uma conta de e-mail ou uma conta do Office 365. Você pode [participar do Programa de Desenvolvedores do Office 365 e obter uma assinatura gratuita de 1 ano do Office 365](https://aka.ms/devprogramsignup).
  - Qualquer navegador que ofereça suporte ECMAScript 5.1, HTML5 e CSS3, como o Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 ou uma versão posterior desses navegadores.
  - Familiaridade com a programação em JavaScript e serviços Web.

<a name="components"></a>
## Componentes principais do exemplo
A solução do exemplo contém os seguintes arquivos:

- AttachmentExampleManifest.xml: O arquivo de manifesto do suplemento do Outlook.
- AppRead\Home\Home.html: A interface do usuário HTML do suplemento do Outlook.
- AppRead\Home\Home.js: O arquivo JavaScript que manipula o envio das informações do anexo ao serviço de Anexo remoto incluído neste exemplo.

O projeto AttachmentService define um serviço REST usando a API do WCF. A exportação contém os seguintes arquivos:

- Controllers\AttachmentServiceController.cs: O objeto de serviço que fornece a lógica de negócios ao serviço do exemplo.
- Models\ServiceRequest: O objeto que representa uma solicitação da Web. O conteúdo do objeto é criado a partir de um objeto de solicitação JSON enviado do seu suplemento.
- Models\Attachment.cs: O objeto utilitário que ajuda a desserializar o objeto JSON que é enviado pelo suplemento de e-mail.
- Models\AttachmentDetails.cs: O objeto que representa os detalhes de cada anexo. Ele fornece um objeto .NET Framework que corresponde ao objeto `AttachmentDetails` dos suplementos de e-mail.
- Models\ServiceResponse: O objeto que representa uma resposta do serviço Web. O conteúdo do objeto é serializado em um objeto JSON ao ser enviado de volta ao suplemento.
- Web.config: Vincula o serviço de amostra ao terminal do servidor da Web.



<a name="codedescription"></a>
##Descrição do código

Este exemplo mostra como recuperar anexos de um serviço Web compatível que suporta seu suplemento de e-mail. Por exemplo, você pode criar um serviço que carrega fotos em um site de compartilhamento ou um serviço que armazena documentos em um repositório. O serviço recebe os anexos diretamente do servidor Exchange e não exige que o cliente execute um processamento extra para obter o anexo e enviá-lo ao serviço.

O exemplo possui duas partes. Na primeira parte o aplicativo do e-mail é executado no cliente do e-mail. O suplemento do e-mail é exibido sempre que uma mensagem ou um compromisso for o item ativo. Quando você seleciona o botão **Testar anexos**, o suplemento do e-mail envia detalhes sobre o anexo ao serviço Web que processa a solicitação. O serviço usa as seguintes etapas para processar o anexo:

- Envia uma solicitação de operação do [GetAttachment](http://msdn.microsoft.com/library/aa494316(v=exchg.150).aspx) ao servidor Exchange que hospeda a caixa de correio. O servidor responde enviando o anexo ao cliente. Neste exemplo, o serviço escreve o XML do servidor para rastrear a saída.
- Retorna o número de anexos processados ao aplicativo do e-mail.



<a name="build"></a>
## Criar e depurar ##
**Observação**: O suplemento de e-mail será ativado em qualquer mensagem de e-mail na caixa de entrada do usuário que tenha um ou mais anexos. Você pode facilitar o teste do suplemento enviando uma ou mais mensagens de e-mail para a sua conta de teste antes de executar o suplemento do exemplo.

1. Abra a solução no Visual Studio.
2. Clique com o botão direito do mouse na solução no Solution Explorer. Selecione **Configurar projetos de inicialização**. 
3. Selecione **Propriedades comuns**e escolha **Inicializar Projeto**.
4. Certifique-se de que a **Ação** do projeto **AttachmentExampleService** esteja definida como **iniciar**.
5. Pressione F5 para criar e implementar o suplemento do exemplo.
6. Conecte-se a uma conta do Exchange fornecendo o endereço de e-mail e a senha de um servidor do Exchange 2013.
7. Permitir que o servidor configure a conta de e-mail.
8. No navegador, faça logon com a conta de e-mail digitando o nome e a senha da conta. 
9. Salvar uma mensagem na Caixa de entrada
10. Aguarde até que a barra de suplementos seja exibida sobre a mensagem.
11. Na barra de suplementos, clique em **AttachmentExample**.
12. Quando o suplemento do e-mail for exibido, clique no botão **TestAttachments** para enviar uma solicitação ao servidor Exchange.
13. O servidor responderá com o número de anexos processados para o item. Isso deve ser igual ao número de anexos contidos no item.

<a name="troubleshooting"></a>
##Troubleshooting
Estes são erros comuns que podem ocorrer quando você usa o Outlook Web App para testar um suplemento de e-mail do Outlook:

- A barra de suplementos não será exibida quando uma mensagem for selecionada. Se isso ocorrer, reinicie o suplemento selecionando **Depurar - Parar a depuração** na janela do Visual Studio e, em seguida, pressione F5 para recriar e implementar o suplemento. 
- Pode ser que as alterações no código JavaScript não sejam selecionadas quando você implementar e executar o suplemento.  Exclua os arquivos temporários da Internet e reinicie o suplemento. 

<a name="questions"></a>
##Perguntas e comentários##

- Se você tiver problemas para executar esse exemplo, [relate um problema](https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments/issues).
- Em geral, perguntas sobre o desenvolvimento de Suplementos do Office devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Não deixe de marcar as perguntas ou comentários com [office-addins].


<a name="additional-resources"></a>
## Recursos adicionais ##

- [Mais exemplos de Suplementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
-  O site oficial do Microsoft ASP.NET](http://www.asp.net/web-api)
- [Tutorial: Obter anexos de um servidor Exchange](http://msdn.microsoft.com/library/dn148008.aspx)

## Direitos autorais
Copyright © 2015 Microsoft. Todos os direitos reservados.


Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
