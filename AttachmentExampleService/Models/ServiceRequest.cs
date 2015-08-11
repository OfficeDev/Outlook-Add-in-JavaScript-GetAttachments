
namespace AttachmentsService.Models
{
  public class ServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public AttachmentDetails[] attachments { get; set; }
  }
}