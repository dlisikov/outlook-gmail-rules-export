using System.Text;

namespace OutlookRulesExport
{
    public class ParsedRule
    {
        public bool MeInCC { get; set; }

        public string ToAddress { get; set; }

        public string FromAddress { get; set; }

        public string SubjectContains { get; set; }

        public string BodyContains { get; set; }

        public string BodyOrSubjectContains { get; set; }

        public string MoveToFolder { get; set; }
        
        public string CopyToFolder { get; set; }
        
        public bool MoveToTrash { get; set; }
        
        public bool DeletePermanently { get; set; }
        
        public bool MeDirectlyOrInCC { get; set; }
        
        public string SenderAddressContains { get; set; }
        
        public string RecipientAddressContains { get; set; }

        public override string ToString()
        {
            var stringBuilder = new StringBuilder();

            stringBuilder.Append("If an email");

            if (!string.IsNullOrEmpty(FromAddress))
            {
                stringBuilder.AppendFormat(" from [{0}]", FromAddress);
            }

            if (!string.IsNullOrEmpty(ToAddress))
            {
                stringBuilder.AppendFormat(" sent to [{0}]", ToAddress);
            }

            if (MeInCC)
            {
                stringBuilder.AppendFormat(" sent to me in CC");
            }

            if (MeDirectlyOrInCC)
            {
                stringBuilder.AppendFormat(" sent to me directly or in CC");
            }
            
            if (!string.IsNullOrEmpty(SenderAddressContains))
            {
                stringBuilder.AppendFormat(" with text in sender addresses [{0}]", SenderAddressContains);
            }
            
            if (!string.IsNullOrEmpty(RecipientAddressContains))
            {
                stringBuilder.AppendFormat(" with text in recipient addresses [{0}]", RecipientAddressContains);
            }

            if (!string.IsNullOrEmpty(SubjectContains))
            {
                stringBuilder.AppendFormat(" with text in subject [{0}]", SubjectContains);
            }

            if (!string.IsNullOrEmpty(BodyContains))
            {
                stringBuilder.AppendFormat(" with text in body [{0}]", BodyContains);
            }

            if (!string.IsNullOrEmpty(BodyOrSubjectContains))
            {
                stringBuilder.AppendFormat(" with text in subject or body [{0}]", BodyOrSubjectContains);
            }

            if (!string.IsNullOrEmpty(MoveToFolder))
            {
                stringBuilder.AppendFormat(" then move to [{0}]", MoveToFolder);
            }

            if (!string.IsNullOrEmpty(CopyToFolder))
            {
                stringBuilder.AppendFormat(" then copy to [{0}]", CopyToFolder);
            }

            if (MoveToTrash)
            {
                stringBuilder.AppendFormat(" then move to trash");
            }
            
            if (DeletePermanently)
            {
                stringBuilder.AppendFormat(" then delete permanently");
            }

            return stringBuilder.ToString();
        }
    }
}