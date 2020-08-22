using ExcelDna.Integration;
using ExcelDna.Registration;
using System;
using System.Linq.Expressions;

namespace StockScraper
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            RegisterFunctions();
        }

        public void AutoClose()
        {
        }

        public void RegisterFunctions()
        {
            var postAsyncReturnConfig = GetPostAsyncReturnConversionConfig();
            ExcelRegistration.GetExcelFunctions()
                .ProcessParameterConversions(postAsyncReturnConfig)
                .ProcessAsyncRegistrations(nativeAsyncIfAvailable: true)
                .RegisterFunctions();
        }

        static ParameterConversionConfiguration GetPostAsyncReturnConversionConfig()
        {
            // This conversion replaces the default #N/A return value of async functions with the #GETTING_DATA value.
            var rval = ExcelError.ExcelErrorGettingData;
            return new ParameterConversionConfiguration()
                .AddReturnConversion((type, customAttributes) => type != typeof(object) ? null : ((Expression<Func<object, object>>)
                                                ((object returnValue) => returnValue.Equals(ExcelError.ExcelErrorNA) ? rval : returnValue)));
        }

        public ExcelFunctionRegistration UpdateHelpTopic(ExcelFunctionRegistration funcReg)
        {
            funcReg.FunctionAttribute.HelpTopic = "http://www.bing.com";
            return funcReg;
        }
    }
}
