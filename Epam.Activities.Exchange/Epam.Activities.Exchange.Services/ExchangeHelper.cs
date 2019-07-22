// License placeholder

using System;
using System.Net;
using System.Security;
using Microsoft.Exchange.WebServices.Data;

namespace Epam.Activities.Exchange.Services
{
    /// <summary>
    /// Exchange service helper.
    /// </summary>
    public static class ExchangeHelper
    {
        /// <summary>
        /// Initializes instance of <see cref="ExchangeService"/> and auto discover url if needed.
        /// </summary>
        /// <param name="login">Master account email.</param>
        /// <param name="password">Master account password.</param>
        /// <param name="exchangeUrl">Exchange service url.</param>
        /// <returns>Instance of <see cref="ExchangeService"/></returns>
        public static ExchangeService GetService(string login, string password, string exchangeUrl)
        {
            var service = new ExchangeService
            {
                Credentials = new NetworkCredential(login, password)
            };

            if (!string.IsNullOrWhiteSpace(exchangeUrl))
            {
                service.Url = new Uri(exchangeUrl);
            }
            else
            {
                service.AutodiscoverUrl(login, AdAutoDiscoCallBack);
            }

            return service;
        }

        /// <summary>
        /// Initializes instance of <see cref="ExchangeService"/> and auto discover url if needed.
        /// </summary>
        /// <param name="login">Master account email.</param>
        /// <param name="password">Master account password.</param>
        /// <param name="exchangeUrl">Exchange service url.</param>
        /// <returns>Instance of <see cref="ExchangeService"/></returns>
        public static ExchangeService GetService(string login, SecureString password, string exchangeUrl)
        {
            var service = new ExchangeService
            {
                Credentials = new NetworkCredential(login, password)
            };

            if (!string.IsNullOrWhiteSpace(exchangeUrl))
            {
                service.Url = new Uri(exchangeUrl);
            }
            else
            {
                service.AutodiscoverUrl(login, AdAutoDiscoCallBack);
            }

            return service;
        }

        /// <summary>
        /// Allows to prevent <see cref="AutodiscoverLocalException"/>
        /// </summary>
        /// <param name="redirectionUrl">Redirection Url</param>
        /// <returns>True for https</returns>
        internal static bool AdAutoDiscoCallBack(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            var result = false;

            var redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }

            return result;
        }
    }
}
