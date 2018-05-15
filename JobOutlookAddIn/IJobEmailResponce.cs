using Utilities;

namespace JobOutlookAddIn
{
	internal interface IJobEmailResponce
	{
		void ImmediateReply( object item, SendCondition condition );
	}
}