/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

using AttachmentsService.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Http;
using System.Xml;

namespace AttachmentsService.Controllers
{
  public class AttachmentServiceController : ApiController
  {
    public ServiceResponse PostAttachments(ServiceRequest request)
    {
      ServiceResponse response = new ServiceResponse();

      try
      {
        response = GetAttachmentsFromExchangeServer(request);
      }
      catch (Exception ex)
      {
        response.isError = true;
        response.message = ex.Message;
      }

      return response;
    }

    // This method does the work of making an Exchange Web Services (EWS) request to get the 
    // attachments from the Exchange server. This implementation makes an individual
    // request for each attachment, and returns the count of attachments processed.
    private ServiceResponse GetAttachmentsFromExchangeServer(ServiceRequest request)
    {
      int processedCount = 0;
      List<string> attachmentNames = new List<string>();

      foreach (AttachmentDetails attachment in request.attachments)
      {
        // Prepare a web request object.
        HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
        webRequest.Headers.Add("Authorization", string.Format("Bearer {0}", request.attachmentToken));
        webRequest.PreAuthenticate = true;
        webRequest.AllowAutoRedirect = false;
        webRequest.Method = "POST";
        webRequest.ContentType = "text/xml; charset=utf-8";

        // Construct the SOAP message for the GetAttchment operation.
        byte[] bodyBytes = Encoding.UTF8.GetBytes(string.Format(GetAttachmentSoapRequest, attachment.id));
        webRequest.ContentLength = bodyBytes.Length;

        Stream requestStream = webRequest.GetRequestStream();
        requestStream.Write(bodyBytes, 0, bodyBytes.Length);
        requestStream.Close();

        // Make the request to the Exchange server and get the response.
        HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

        // If the response is okay, create an XML document from the
        // response and process the request.
        if (webResponse.StatusCode == HttpStatusCode.OK)
        {
          Stream responseStream = webResponse.GetResponseStream();

          XmlDocument xmlDocument = new XmlDocument();
          xmlDocument.Load(responseStream);

          // This method simply writes the XML document to the
          // trace output. Your service would perform its
          // processing here.
          Trace.Write(xmlDocument.InnerXml);

          // Close the response stream.
          responseStream.Close();
          webResponse.Close();

          processedCount++;
          attachmentNames.Add(attachment.name);
        }

      }
      ServiceResponse response = new ServiceResponse();
      response.attachmentNames = attachmentNames.ToArray();
      response.attachmentsProcessed = processedCount;

      return response;
    }

    private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
  }
}

// *********************************************************
//
// Outlook-Add-in-Javascript-GetAttachments, https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************