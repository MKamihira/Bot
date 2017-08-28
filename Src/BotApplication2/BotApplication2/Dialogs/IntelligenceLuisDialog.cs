using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System.Collections.ObjectModel;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using outlook = Microsoft.Office.Interop.Outlook;

namespace BotApplication2.Dialogs
{
    [LuisModel("a9410e1a-245f-4bbe-8772-e19f6da0b277", "403554440d484dc59beb5896b11e4141")]
    [Serializable]
    public class IntelligenceLuisDialog : LuisDialog<object>
    {
        [LuisIntent("会議予定作りたい")]
        public async Task Meeting(IDialogContext context, LuisResult result)
        {
            string date = "";
            string time = "";
            string locate = "";
            string theme = "";
            DateTime? time1 = null;
            DateTime? time2 = null;

            foreach (EntityRecommendation entity in result.Entities)
            {
                if (entity.Type == "日付")
                {
                    date = entity.Entity;
                }
                else if (entity.Type == "時間")
                {
                    System.Globalization.CultureInfo ci =
                        new System.Globalization.CultureInfo("ja-JP");
                    if (time1 == null)
                    {
                        time1 = DateTime.Parse(entity.Entity, ci,
                            System.Globalization.DateTimeStyles.AssumeLocal);
                    }
                    else
                    {
                        time2 = DateTime.Parse(entity.Entity, ci,
                            System.Globalization.DateTimeStyles.AssumeLocal);
                    }
                }
                else if (entity.Type == "場所")
                {
                    locate = entity.Entity;
                }
                else if (entity.Type == "テーマ")
                {
                    theme = entity.Entity;
                }
            }
            if (time2 == null && time1 != null)
            {
                time = time1.ToString().Substring(11, 5) + " ~";
            }
            else if (time2 != null)
            {
                if (time1 < time2)
                {
                    time = time1.ToString().Substring(11, 5) + " ~" + time2.ToString().Substring(11, 5);
                }
                else
                {
                    time = time2.ToString().Substring(11, 5) + " ~" + time1.ToString().Substring(11, 5);
                }
            }

            await context.PostAsync("日にち ： " + date);
            await context.PostAsync("時間 : " + time);
            await context.PostAsync("場所 : " + locate);
            await context.PostAsync("テーマ : " + theme);
            await context.PostAsync("会議出席依頼送信画面を開きます");

            try
            {
                var app = new outlook.Application();

                outlook.AppointmentItem appt =
                    app.CreateItem(outlook.OlItemType.olAppointmentItem)
                    as outlook.AppointmentItem;
                appt.MeetingStatus = outlook.OlMeetingStatus.olMeeting;
                if (time1 != null)
                appt.Start = DateTime.Parse(time1.ToString());
                if (time2 != null)
                appt.End = DateTime.Parse(time2.ToString());
                appt.Location = locate;
                //appt.Body = body.Text;
                appt.AllDayEvent = false;
                appt.Subject = theme;
                //appt.Recipients.Add("MKamihira@mx1.wiseman.co.jp");
                //outlook.Recipients sentTo = appt.Recipients;
                //outlook.Recipient sentInvite = null;
                //string[] ad = sentto.Text.Split(':');
                //sentInvite = sentTo.Add(ad[1]);
                //sentInvite.Type = (int)outlook.OlMeetingRecipientType.olRequired;
                //sentTo.ResolveAll();
                appt.Save();
                appt.Display();
                //appt.Send();
            }
            catch (Exception ex)
            {
                await context.PostAsync($"Failed with message:{ex.Message}");
            }

            context.Wait(this.MessageReceived);
        }

        [LuisIntent("挨拶したい")]
        public async Task Greeting(IDialogContext context, LuisResult result)
        {
            string message = "";
            foreach (EntityRecommendation entity in result.Entities)
            {
                if (entity.Type == "挨拶")
                {
                    message = entity.Entity;
                }
                else if (entity.Type == "名前")
                {
                    message = message + entity.Entity + "マン";
                }
            }
            await context.PostAsync(message);

            context.Wait(this.MessageReceived);
        }

        [LuisIntent("")]
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            string message = "ばかか？";

            await context.PostAsync(message);

            context.Wait(this.MessageReceived);
        }
    }
}