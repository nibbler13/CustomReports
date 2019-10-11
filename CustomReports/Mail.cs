using System.Net.Mail;
using System;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Net.Mime;
using System.Threading.Tasks;
using System.Diagnostics;

namespace CustomReports {
	public static class Mail {
		public static void SendMail (string subject, string body, string receiver, string attachmentPath = "") {
			Logging.ToLog("Mail - Отправка сообщения, тема: " + subject + ", текст: " + body);
			Logging.ToLog("Mail - Получатели: " + receiver);

			if (string.IsNullOrEmpty(receiver)) {
				Logging.ToLog("Mail - Пропуск отправки, не задан получатель");
				return;
			}

			try { 
				SmtpClient client = CreateClientAndMessage(subject, body, receiver, out MailMessage message, attachmentPath);
				client.Send(message);
				Logging.ToLog("Mail - Письмо отправлено успешно");
				DisposeResources(client, message);
			} catch (Exception e) {
				Logging.ToLog("Mail - SendMail exception: " + e.Message + Environment.NewLine + e.StackTrace);
			}
		}

		public static void SendTestMail(string subject, string body, string receiver) {
			SmtpClient client = CreateClientAndMessage(subject, body, receiver, out MailMessage message);
			client.Send(message);
			DisposeResources(client, message);
		}

		private static SmtpClient CreateClientAndMessage(string subject, string body, string receiver, out MailMessage message, string attachmentPath = "") {
			string appName = Assembly.GetExecutingAssembly().GetName().Name;
			if (!string.IsNullOrEmpty(Configuration.Instance.MailSenderName))
				appName = Configuration.Instance.MailSenderName;

			MailAddress from = new MailAddress(Configuration.Instance.MailUser, appName);
			List<MailAddress> mailAddressesTo = new List<MailAddress>();

			if (receiver.Contains(" | ")) {
				string[] receivers = Configuration.GetSplittedAddresses(receiver);
				foreach (string address in receivers)
					try {
						mailAddressesTo.Add(new MailAddress(address));
					} catch (Exception e) {
						Logging.ToLog("Mail - Не удалось разобрать адрес: " + address + Environment.NewLine + e.Message);
					}
			} else
				try {
					mailAddressesTo.Add(new MailAddress(receiver));
				} catch (Exception e) {
					Logging.ToLog("Mail - Не удалось разобрать адрес: " + receiver + Environment.NewLine + e.Message);
				}

			if (!string.IsNullOrEmpty(Configuration.Instance.MailSign))
				body += Environment.NewLine + Environment.NewLine +
					Configuration.Instance.MailSign.Replace("@machineName", Environment.MachineName);

			message = new MailMessage();

			foreach (MailAddress mailAddress in mailAddressesTo)
				message.To.Add(mailAddress);

			message.IsBodyHtml = body.Contains("<") && body.Contains(">");

			if (message.IsBodyHtml)
				body = body.Replace(Environment.NewLine, "<br>");

			if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath)) {
#pragma warning disable IDE0068 // Use recommended dispose pattern
				Attachment attachment = new Attachment(attachmentPath);
#pragma warning restore IDE0068 // Use recommended dispose pattern

				if (message.IsBodyHtml && attachmentPath.EndsWith(".jpg")) {
					attachment.ContentDisposition.Inline = true;

					LinkedResource inline = new LinkedResource(attachmentPath, MediaTypeNames.Image.Jpeg) {
						ContentId = Guid.NewGuid().ToString()
					};

					body = body.Replace("Фотография с камеры терминала:", "Фотография с камеры терминала:<br>" +
						string.Format(@"<img src=""cid:{0}"" />", inline.ContentId));

					AlternateView avHtml = AlternateView.CreateAlternateViewFromString(body, null, MediaTypeNames.Text.Html);
					avHtml.LinkedResources.Add(inline);

					message.AlternateViews.Add(avHtml);
				} else
					message.Attachments.Add(attachment);
			}

			message.From = from;
			message.Subject = subject;
			message.Body = body;

			if (CustomReports.Configuration.Instance.ShouldAddAdminToCopy) {
				string adminAddress = CustomReports.Configuration.Instance.MailAdminAddress;
				if (!string.IsNullOrEmpty(adminAddress))
					if (adminAddress.Contains(" | ")) {
						string[] adminAddresses = CustomReports.Configuration.GetSplittedAddresses(adminAddress);
						foreach (string address in adminAddresses)
							try {
								message.CC.Add(new MailAddress(address));
							} catch (Exception e) {
								Logging.ToLog("Mail - Не удалось разобрать адрес: " + address + Environment.NewLine + e.Message);
							}
					} else
						try {
							message.CC.Add(new MailAddress(adminAddress));
						} catch (Exception e) {
							Logging.ToLog("Mail - Не удалось разобрать адрес: " + adminAddress + Environment.NewLine + e.Message);
						}
			}

			SmtpClient client = new SmtpClient(Configuration.Instance.MailSmtpServer, (int)Configuration.Instance.MailSmtpPort) {
				UseDefaultCredentials = false,
				DeliveryMethod = SmtpDeliveryMethod.Network,
				EnableSsl = Configuration.Instance.MailEnableSSL,
				Credentials = new System.Net.NetworkCredential(
				Configuration.Instance.MailUser,
				Configuration.Instance.MailPassword)
			};

			if (!string.IsNullOrEmpty(Configuration.Instance.MailUserDomain))
				client.Credentials = new System.Net.NetworkCredential(
				Configuration.Instance.MailUser,
				Configuration.Instance.MailPassword,
				Configuration.Instance.MailUserDomain);

			return client;
		}

		private static void DisposeResources(SmtpClient client, MailMessage message) {
			client.Dispose();
			foreach (Attachment attach in message.Attachments)
				attach.Dispose();

			message.Dispose();
		}
	}
}
