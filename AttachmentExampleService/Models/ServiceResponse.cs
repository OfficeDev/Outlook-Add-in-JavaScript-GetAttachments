
namespace AttachmentsService.Models
{
  public class ServiceResponse
  {
    public bool isError { get; set; }
    public string message { get; set; }
    public int attachmentsProcessed { get; set; }
    public string[] attachmentNames { get; set; }
  }
}