using FlowSynx.PluginCore;

namespace FlowSynx.Plugins.Email.Receiver.Models;

public class EmailReceiverPluginSpecifications : PluginSpecifications
{
    [RequiredMember]
    public string Host { get; set; } = string.Empty;

    public int Port { get; set; } = 993;

    public bool UseSsl { get; set; } = true;
    
    [RequiredMember]
    public string Username { get; set; } = string.Empty;
    
    [RequiredMember]
    public string Password { get; set; } = string.Empty;

    [RequiredMember]
    public string From { get; set; } = string.Empty;
}