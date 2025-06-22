using FlowSynx.PluginCore.Helpers;
using FlowSynx.PluginCore;
using FlowSynx.PluginCore.Extensions;
using MailKit.Security;
using FlowSynx.Plugins.Email.Receiver.Models;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using MailKit;

namespace FlowSynx.Plugins.Email.Receiver;

public class EmailReceiverPlugin : IPlugin
{
    private IPluginLogger? _logger;
    private EmailReceiverPluginSpecifications _emailReceiverSpecifications = null!;
    private bool _isInitialized;

    public PluginMetadata Metadata
    {
        get
        {
            return new PluginMetadata
            {
                Id = Guid.Parse("32d20661-28f0-493b-975b-9307d5cf3d5b"),
                Name = "Email.Receiver",
                CompanyName = "FlowSynx",
                Description = Resources.PluginDescription,
                Version = new PluginVersion(1, 0, 0),
                Namespace = PluginNamespace.Connectors,
                Authors = new List<string> { "FlowSynx" },
                Copyright = "© FlowSynx. All rights reserved.",
                Icon = "flowsynx.png",
                ReadMe = "README.md",
                RepositoryUrl = "https://github.com/flowsynx/plugin-email-receiver",
                ProjectUrl = "https://flowsynx.io",
                Tags = new List<string>() { "flowSynx", "email", "email-receiver", "communication", "collaboration" },
                Category = PluginCategories.Communication
            };
        }
    }

    public PluginSpecifications? Specifications { get; set; }

    public Type SpecificationsType => typeof(EmailReceiverPluginSpecifications);

    public IReadOnlyCollection<string> SupportedOperations => new List<string>();

    public Task Initialize(IPluginLogger logger)
    {
        if (ReflectionHelper.IsCalledViaReflection())
            throw new InvalidOperationException(Resources.ReflectionBasedAccessIsNotAllowed);

        ArgumentNullException.ThrowIfNull(logger);
        _emailReceiverSpecifications = Specifications.ToObject<EmailReceiverPluginSpecifications>();
        _logger = logger;
        _isInitialized = true;
        return Task.CompletedTask;
    }

    public async Task<object?> ExecuteAsync(PluginParameters parameters, CancellationToken cancellationToken)
    {
        EnsureValidCall();

        var emailParams = parameters.ToObject<EmailReceiverMessageParameters>();
        using var imapClient = await ConnectAndAuthenticateAsync(cancellationToken);

        var inbox = imapClient.Inbox;
        await inbox.OpenAsync(MailKit.FolderAccess.ReadOnly, cancellationToken);

        var uids = await inbox.SearchAsync(SearchQuery.NotSeen, cancellationToken);
        var messages = await FetchMessagesAsync(inbox, uids, emailParams.MaxResults ?? 10, cancellationToken);

        await imapClient.DisconnectAsync(true, cancellationToken);

        _logger?.LogInfo($"Retrieved {messages.Count} emails.");
        return new { count = messages.Count, messages };
    }

    private void EnsureValidCall()
    {
        if (ReflectionHelper.IsCalledViaReflection())
            throw new InvalidOperationException(Resources.ReflectionBasedAccessIsNotAllowed);

        if (!_isInitialized)
            throw new InvalidOperationException($"Plugin '{Metadata.Name}' v{Metadata.Version} is not initialized.");
    }

    private async Task<ImapClient> ConnectAndAuthenticateAsync(CancellationToken cancellationToken)
    {
        var imapClient = new ImapClient();
        await imapClient.ConnectAsync(
            _emailReceiverSpecifications.Host,
            _emailReceiverSpecifications.Port,
            _emailReceiverSpecifications.UseSsl ? SecureSocketOptions.StartTls : SecureSocketOptions.Auto,
            cancellationToken);

        await imapClient.AuthenticateAsync(_emailReceiverSpecifications.Username, _emailReceiverSpecifications.Password, cancellationToken);
        return imapClient;
    }

    private async Task<List<PluginContext>> FetchMessagesAsync(IMailFolder inbox, IList<UniqueId> uids, int maxResults, CancellationToken cancellationToken)
    {
        var messages = new List<PluginContext>();

        foreach (var uid in uids.Take(maxResults))
        {
            var message = await inbox.GetMessageAsync(uid, cancellationToken);
            messages.Add(CreatePluginContextFromMessage(message));
        }

        return messages;
    }

    private PluginContext CreatePluginContextFromMessage(MimeMessage message)
    {
        var id = Guid.NewGuid().ToString();
        var context = new PluginContext(id, "Email")
        {
            Format = message.HtmlBody != null ? "HTML" : "Text",
            Content = message.TextBody ?? message.HtmlBody
        };

        context.Metadata["From"] = message.From.ToString();
        context.Metadata["Subject"] = message.Subject;
        context.Metadata["Date"] = message.Date.ToString("u");

        var attachments = ExtractAttachments(message);
        if (attachments.Any())
        {
            context.Metadata["HasAttachments"] = true;
            context.Metadata["Attachments"] = attachments;
        }
        else
        {
            context.Metadata["HasAttachments"] = false;
        }

        return context;
    }

    private List<Dictionary<string, object>> ExtractAttachments(MimeMessage message)
    {
        var attachments = new List<Dictionary<string, object>>();

        foreach (var attachment in message.Attachments)
        {
            if (attachment is MimePart part)
            {
                using var stream = new MemoryStream();
                part.Content.DecodeTo(stream);

                attachments.Add(new Dictionary<string, object>
                {
                    ["FileName"] = part.FileName,
                    ["ContentType"] = part.ContentType.MimeType,
                    ["Size"] = stream.Length,
                    ["Data"] = stream.ToArray()
                });
            }
            else if (attachment is MessagePart rfc822Part)
            {
                using var stream = new MemoryStream();
                rfc822Part.Message.WriteTo(stream);

                attachments.Add(new Dictionary<string, object>
                {
                    ["FileName"] = GetMessagePartFileName(rfc822Part),
                    ["ContentType"] = "message/rfc822",
                    ["Size"] = stream.Length,
                    ["Data"] = stream.ToArray()
                });
            }
        }

        return attachments;
    }

    private string GetMessagePartFileName(MessagePart messagePart)
    {
        return !string.IsNullOrEmpty(messagePart.ContentDisposition?.FileName)
            ? messagePart.ContentDisposition.FileName
            : !string.IsNullOrEmpty(messagePart.ContentType?.Name)
                ? messagePart.ContentType.Name
                : "message.eml";
    }
}