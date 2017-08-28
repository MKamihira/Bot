using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System.Collections.ObjectModel;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using System.Threading;

namespace BotApplication2.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        //private string luisUrl = @"https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/a9410e1a-245f-4bbe-8772-e19f6da0b277?subscription-key=403554440d484dc59beb5896b11e4141&verbose=true&timezoneOffset=0&q={0}";

        public Task StartAsync(IDialogContext context)
        {
            //context.PostAsync("こんにちは");

            context.Wait(MessageReceivedAsync);

            return Task.CompletedTask;
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<object> result)
        {
            var activity = await result as Activity;

            await context.Forward(new IntelligenceLuisDialog(), ResumeAfterOptionDialog, activity, CancellationToken.None);
        }

        private async Task ResumeAfterOptionDialog(IDialogContext context, IAwaitable<object> result)
        {
            try
            {
                var message = await result;
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Failed with message:{ex.Message}");
            }
            finally
            {
                context.Wait(this.MessageReceivedAsync);
            }
        }
    }
}