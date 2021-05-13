---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
description: Cet exemple illustre comment récupérer des pièces jointes depuis une boîte aux lettres Exchange.
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/11/2015 1:48:02 PM
---
# Complément Outlook : Obtenir des pièces jointes depuis un serveur Exchange

**Table des matières**

* [Résumé](#summary)
* [Conditions préalables](#prerequisites)
* [Composants clés de l’exemple](#components)
* [Description du code](#codedescription)
* [Création et débogage](#build)
* [Résolution des problèmes](#troubleshooting)
* [Questions et commentaires](#questions)
* [Ressources supplémentaires](#additional-resources)

<a name="summary"></a>
##Cet exemple
illustre comment obtenir des pièces jointes depuis une boîte aux lettres Exchange.

<a name="prerequisites"></a>
## Conditions préalables ##

Cet exemple nécessite les éléments suivants :  

  - Visual Studio 2013 avec la mise à jour 5 ou Visual Studio 2015.  
  - Un ordinateur exécutant Exchange 2013 avec au moins un compte de messagerie ou un compte Office 365. Vous pouvez [participer au programme pour les développeurs Office 365 et obtenir un abonnement gratuit d’un an à Office 365](https://aka.ms/devprogramsignup).
  - Tout navigateur qui prend en charge ECMAScript 5.1, HTML5 et CSS3, tel qu’Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 ou une version ultérieure de ces navigateurs.
  - Être familiarisé avec les services web et de programmation JavaScript.

<a name="components"></a>
## Composants clés de l’exemple
La solution de l’exemple contient les fichiers suivants :

- AttachmentExampleManifest.xml : Fichier manifeste pour le complément Outlook.
- AppRead\Home\Home.html : Interface utilisateur HTML pour le complément messagerie pour Outlook.
- AppRead\Home\Home.js : Fichier JavaScript qui gère l’envoi des informations de la pièce jointe au service distant de pièces jointes inclus dans cet exemple.

Le projet AttachmentService définit un service REST à l’aide de l’API WCF. Le projet contient les fichiers suivants :

- Controllers\AttachmentServiceController.cs : Objet de service qui fournit la logique métier pour l’exemple de service.
- Models\ServiceRequest : Objet qui représente une requête web. Le contenu de l’objet est créé à partir d’un objet de requête JSON envoyé depuis votre complément messagerie.
- Models\Attachment.cs : Objet utilitaire qui permet de désérialiser l’objet JSON envoyé par le complément messagerie.
- Models\AttachmentDetails.cs : Objet qui représente les détails de chaque pièce jointe. Fournit un objet .NET Framework qui correspond à l’objet `AttachmentDetails` du complément messagerie.
- Models\ServiceResponse : Objet qui représente une réponse du service Web. Le contenu de l’objet est sérialisé en objet JSON lorsqu’il est renvoyé au complément messagerie.
- Web.config : Lie l’exemple de service au point de terminaison serveur Web.



<a name="codedescription"></a>
##Description du code

Cet exemple illustre comment récupérer des pièces jointes depuis un service Web qui prend en charge votre complément messagerie. Par exemple, vous pouvez créer un service qui télécharge des photos vers un site de partage, ou un service qui stocke des documents dans un référentiel. Le service obtient les pièces jointes directement du serveur Exchange Server sans obliger le client à effectuer un traitement supplémentaire pour récupérer la pièce jointe, puis l’envoie en même temps que le service.

L’exemple comporte deux parties : La première, l’application courrier, s’exécute dans le client de messagerie. Le complément messagerie s’affiche chaque fois qu’un message ou un rendez-vous est l’élément actif. Lorsque vous sélectionnez le bouton **Tester les pièces jointes**, le complément messagerie envoie des détails sur la pièce jointe au service web qui traite la demande. Le service suit les étapes suivantes pour traiter les pièces jointes :

- Envoi d’une demande d’opération [GetAttachment](http://msdn.microsoft.com/library/aa494316(v=exchg.150).aspx) au serveur Exchange qui héberge la boîte aux lettres. Le serveur répond en envoyant la pièce jointe au service. Dans cet exemple, le service écrit simplement le code XML du serveur pour suivre la sortie.
- Renvoi du nombre de pièces jointes traitées à l’application courrier.



<a name="build"></a>
## Création et débogage ##
**Remarque** : Le complément messagerie sera activé sur tout message électronique figurant dans la boîte de réception de l’utilisateur et contenant au moins une pièce jointe. Vous pouvez simplifier le test du complément en envoyant un ou plusieurs courriers électroniques à votre compte de test avant d’exécuter l’exemple de complément.

1. Ouvrez la solution dans Visual Studio.
2. Cliquez droit sur la solution dans l’Explorateur de solutions. Sélectionnez **Définir des projets de démarrage**. 
3. Sélectionnez **Propriétés communes**, puis sélectionnez **Projets de démarrage**.
4. Assurez-vous que l’**Action** pour le projet **AttachmentExampleService** est réglée sur **Démarrer**.
5. Appuyez sur F5 pour créer et déployer l’exemple de complément.
6. Connectez-vous à un compte Exchange en fournissant l’adresse de courrier et le mot de passe d’un serveur Exchange 2013.
7. Autorisez le serveur à configurer le compte de messagerie.
8. Connectez-vous au compte de courrier en entrant le nom du compte et le mot de passe. 
9. Sélectionnez un message dans la boîte de réception.
10. Patientez jusqu’à ce que la barre du complément s’affiche au-dessus du message.
11. Dans la barre du complément, cliquez sur **AttachmentExample**.
12. Lorsque le complément messagerie apparaît, cliquez sur le bouton **TestAttachments** pour envoyer une demande au serveur Exchange.
13. Le serveur répond avec le nombre de pièces jointes traitées pour l’élément. Ce nombre doit être égal à celui des pièces jointes contenues dans l’élément.

<a name="troubleshooting"></a>
##Dépannage
Voici quelques erreurs courantes qui peuvent se produire lors de l’utilisation d’Outlook Web App pour tester un complément messagerie pour Outlook :

- La barre du complément n’apparaît pas lorsqu’un message est sélectionné. Si c’est le cas, redémarrez l’application en sélectionnant **Debug – arrêter le débogage** dans la fenêtre Visual Studio, puis appuyez sur F5 pour regénérer et déployer le complément. 
- Les modifications apportées au code JavaScript peuvent ne pas être prises en compte lors du déploiement et de l’exécution du complément. Si les modifications ne sont pas prises en compte, effacez le cache du navigateur web en sélectionnant **Outils – Options Internet** puis en cliquant sur le bouton **Supprimer...** Supprimez les fichiers Internet temporaires, puis redémarrez le complément. 

<a name="questions"></a>
##Questions et commentaires##

- Si vous rencontrez des difficultés à l’exécution de cet exemple, veuillez [consigner un problème](https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments/issues).
- Si vous avez des questions générales sur le développement de compléments Office, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Posez vos questions ou envoyez vos commentaires en incluant la balise [office-addins].


<a name="additional-resources"></a>
## Ressources supplémentaires ##

- [Autres exemples de compléments](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [API Web : Le site officiel Microsoft ASP.NET](http://www.asp.net/web-api)
- [Procédure : Obtenir des pièces jointes depuis un serveur Exchange](http://msdn.microsoft.com/library/dn148008.aspx)

## Copyright
Copyright (c) 2015 Microsoft. Tous droits réservés.


Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
