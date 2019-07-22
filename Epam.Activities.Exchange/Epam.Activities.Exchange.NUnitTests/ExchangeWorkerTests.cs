// License placeholder

using System;
using System.Collections.Generic;
using System.IO;
using AutoFixture;
using Epam.Activities.Exchange.Data.Interfaces;
using Epam.Activities.Exchange.Services;
using Microsoft.Exchange.WebServices.Data;
using Moq;
using NUnit.Framework;

namespace Epam.Activities.Exchange.NUnitTests
{
    [TestFixture]
    public class ExchangeWorkerTests
    {
        private ExchangeWorker _target;
        private Mock<IEmailMessageBuilder> _emailMessage;

        [SetUp]
        public void Init()
        {
            Logger.Init(null, null, null, null, true);
            _emailMessage = new Mock<IEmailMessageBuilder>();
            _target = new ExchangeWorker(_emailMessage.Object);
        }
        
        [Test]
        public void SendEmailMessage_Positive()
        {
            _emailMessage.Setup(e => e.EmailMessage).Returns(new EmailMessage(new ExchangeService()));
            _emailMessage.Setup(e => e.Send());

            Assert.DoesNotThrow(() => _target.SendEmailMesage(
                new List<string>(),
                new List<string>(),
                new Fixture().Create<string>(),
                new Fixture().Create<string>(),
                new List<string>()));

            _emailMessage.Verify(e => e.EmailMessage, Times.Exactly(1));
            _emailMessage.Verify(e => e.Send(), Times.Exactly(1));            
        }

        [Test]
        [TestCase(null, null, null, null, null)]
        [TestCase(null, null, "body", null, null)]
        [TestCase(null, null, null, "test", null)]
        [TestCase(null, null, "body", "test", null)]
        public void SendEmailMessage_ArgumentNullException(IList<string> to, IList<string> cc, string body, string subject, IList<string> attachments)
        {
            _emailMessage.Setup(e => e.EmailMessage).Returns(It.IsAny<EmailMessage>());
            _emailMessage.Setup(e => e.Send());

            Assert.Throws<ArgumentNullException>(() => _target.SendEmailMesage(to, cc, body, subject, attachments));

            _emailMessage.Verify(e => e.EmailMessage, Times.Exactly(0));
            _emailMessage.Verify(e => e.Send(), Times.Exactly(0));
        }

        [Test]
        public void SendEmailMessage_SomeException()
        {
            _emailMessage.Setup(e => e.EmailMessage).Returns(new EmailMessage(new ExchangeService()));
            _emailMessage.Setup(e => e.Send()).Throws(new Exception());

            Assert.Throws<Exception>(() => _target.SendEmailMesage(
                new List<string>(),
                new List<string>(),
                null,
                null,
                new List<string>()));

            _emailMessage.Verify(e => e.EmailMessage, Times.Exactly(1));
            _emailMessage.Verify(e => e.Send(), Times.Exactly(1));
        }

        [Test]
        public void SendEmailMessage_InvalidDataException()
        {
            _emailMessage.Setup(e => e.EmailMessage).Returns(new EmailMessage(new ExchangeService()));
            _emailMessage.Setup(e => e.Send());

            Assert.Throws<InvalidDataException>(() => _target.SendEmailMesage(
                new List<string>(),
                new List<string>(),
                null,
                null,
                new List<string>(),
                TestMethod));

            _emailMessage.Verify(e => e.EmailMessage, Times.Exactly(1));
            _emailMessage.Verify(e => e.Send(), Times.Exactly(0));
        }

        [Test]
        public void SendEmailMessageAsync_Positive()
        {
            _emailMessage.Setup(e => e.EmailMessage).Returns(new EmailMessage(new ExchangeService()));
            _emailMessage.Setup(e => e.Send());

            Assert.DoesNotThrowAsync(() => _target.SendEmailMesageAsync(
                new List<string>(),
                new List<string>(),
                new Fixture().Create<string>(),
                new Fixture().Create<string>(),
                new List<string>()));

            _emailMessage.Verify(e => e.EmailMessage, Times.Exactly(1));
            _emailMessage.Verify(e => e.Send(), Times.Exactly(1));
        }

        [Test]
        [TestCase(null, null, null, null, null)]
        [TestCase(null, null, "body", null, null)]
        [TestCase(null, null, null, "test", null)]
        [TestCase(null, null, "body", "test", null)]
        public void SendEmailMessageAsync_ArgumentNullException(IList<string> to, IList<string> cc, string body, string subject, IList<string> attachments)
        {
            _emailMessage.Setup(e => e.EmailMessage).Returns(It.IsAny<EmailMessage>());
            _emailMessage.Setup(e => e.Send());

            Assert.ThrowsAsync<ArgumentNullException>(() => _target.SendEmailMesageAsync(to, cc, body, subject, attachments));

            _emailMessage.Verify(e => e.EmailMessage, Times.Exactly(0));
            _emailMessage.Verify(e => e.Send(), Times.Exactly(0));
        }

        [Test]
        public void SendEmailMessageAsync_SomeException()
        {
            _emailMessage.Setup(e => e.EmailMessage).Returns(new EmailMessage(new ExchangeService()));
            _emailMessage.Setup(e => e.Send()).Throws(new Exception());

            Assert.ThrowsAsync<Exception>(() => _target.SendEmailMesageAsync(
                new List<string>(),
                new List<string>(),
                null,
                null,
                new List<string>()));

            _emailMessage.Verify(e => e.EmailMessage, Times.Exactly(1));
            _emailMessage.Verify(e => e.Send(), Times.Exactly(1));
        }

        [Test]
        public void SendEmailMessageAsync_InvalidDataException()
        {
            _emailMessage.Setup(e => e.EmailMessage).Returns(new EmailMessage(new ExchangeService()));
            _emailMessage.Setup(e => e.Send());

            Assert.ThrowsAsync<InvalidDataException>(() => _target.SendEmailMesageAsync(
                new List<string>(),
                new List<string>(),
                null,
                null,
                new List<string>(),
                TestMethod));

            _emailMessage.Verify(e => e.EmailMessage, Times.Exactly(1));
            _emailMessage.Verify(e => e.Send(), Times.Exactly(0));
        }

        private static IList<EmailAddress> TestMethod(IEnumerable<string> adresses)
        {
            throw new InvalidDataException();
        }
    }
}
