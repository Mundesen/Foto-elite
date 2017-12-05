using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Compare.Enums
{
   public enum MailSentStatus
    {
        BelowAmount,
        ImageNotFound,
        ChosenImageNumberMissing,
        MailSend,
        InProgress,
        ReadyForMail,
        MailNotSent
    }
}
