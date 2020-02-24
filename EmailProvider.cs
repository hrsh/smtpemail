using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Fartak.Toolkit.FartakCustomToolkit.ActionResult;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Extensions.Options;
using MimeKit;

namespace Fartak.Toolkit.FartakCustomToolkit.Email
{
    public class EmailProvider : IEmailProvider, IDisposable
    {
        private readonly FartakSmtpModel _smtpConfig;

        public EmailProvider(IOptionsSnapshot<FartakSmtpModel> configuration)
        {
            var myConfig = configuration;
            if (myConfig == null)
                throw new ArgumentNullException(nameof(myConfig));

            _smtpConfig = myConfig.Value;
            if (_smtpConfig == null)
                throw new ArgumentNullException(nameof(_smtpConfig));
        }

        private MimeMessage PrepairEmailMessage(
            IList<FartakMailAddress> blindCarbonCopies,
            IList<FartakMailAddress> carbonCopies,
            string subject,
            string body,
            string name,
            string address,
            FartakMailHeaders headers = null)
        {
            var mailMessage = new MimeMessage();
            mailMessage.From.Add(new MailboxAddress(_smtpConfig.FromName, _smtpConfig.FromAddress));
            mailMessage.Subject = subject;
            mailMessage.To.Add(new MailboxAddress(name, address));
            if (blindCarbonCopies != null && blindCarbonCopies.Any())
            {
                foreach (var bcc in blindCarbonCopies)
                {
                    mailMessage.Bcc.Add(bcc.PrepairAddress());
                }
            }

            if (carbonCopies != null && carbonCopies.Any())
            {
                foreach (var cc in carbonCopies)
                {
                    mailMessage.Cc.Add(cc.PrepairAddress());
                }
            }

            mailMessage.Body = GetMessageBody(body);
            mailMessage.AddHeader(headers, _smtpConfig.FromAddress);
            return mailMessage;
        }

        private static MimeEntity GetMessageBody(string body)
        {
            var builder = new BodyBuilder {HtmlBody = body};
            return builder.ToMessageBody();
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public async Task<FartakActionResult<FartakEmailModel>> SendAsync(FartakEmailModel model)
        {
            using var client = new SmtpClient();
            if (!string.IsNullOrWhiteSpace(_smtpConfig.LocalDomain))
                client.LocalDomain = _smtpConfig.LocalDomain;

            try
            {
                await client.ConnectAsync(_smtpConfig.Server, _smtpConfig.Port, SecureSocketOptions.None);
            }
            catch (Exception e)
            {
                model.Sent = false;
                return FartakActionResult<FartakEmailModel>.Failed(new FartakResultError
                {
                    Description = e.ToString()
                });
            }
            if (!string.IsNullOrWhiteSpace(_smtpConfig.Username) &&
                !string.IsNullOrWhiteSpace(_smtpConfig.Password))
                await client.AuthenticateAsync(_smtpConfig.Username, _smtpConfig.Password);

            var message = PrepairEmailMessage(model.BlindCarbonCopies, model.CarbonCopies, model.Subject,
                model.Body, model.EmailAddresse.Name, model.EmailAddresse.Address, model.Headers);
            try
            {
                await client.SendAsync(message);
                model.Sent = true;
                model.Message = $"Email sent successfuly on {DateTimeOffset.UtcNow}";
                return FartakActionResult<FartakEmailModel>.Invoke(model);
            }
            catch (Exception e)
            {
                model.Sent = false;
                model.Message = e.ToString();
                return FartakActionResult<FartakEmailModel>.Failed();
            }
        }

        public async Task<FartakActionResult<FartakEmailModel>> SendAsync(string subject, string body, string toName,
            string toAddress)
        {
            var model = new FartakEmailModel
            {
                Body = body,
                Subject = subject,
                EmailAddresse = new FartakMailAddress
                {
                    Name = toName,
                    Address = toAddress
                }
            };
            return await SendAsync(model);
        }

        public async Task<FartakActionResult<FartakEmailModel>> SendToAdminAsync(string body, string subject = "")
        {
            var model = new FartakEmailModel
            {
                Body = body,
                Subject = string.IsNullOrWhiteSpace(subject) ? "ارسال پیام هشدار از فرتاک" : subject,
                EmailAddresse = new FartakMailAddress
                {
                    Name = "HRShojaie",
                    Address = _smtpConfig.AdminAddress
                }
            };
            return await SendAsync(model);
        }
    }
}
