#pragma checksum "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\GenerateFile.cshtml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "53914d767f02b7acc15b4d5f9efb12ff5af27867"
// <auto-generated/>
#pragma warning disable 1591
[assembly: global::Microsoft.AspNetCore.Razor.Hosting.RazorCompiledItemAttribute(typeof(AspNetCore.Views_Diligence_GenerateFile), @"mvc.1.0.view", @"/Views/Diligence/GenerateFile.cshtml")]
namespace AspNetCore
{
    #line hidden
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.AspNetCore.Mvc.Rendering;
    using Microsoft.AspNetCore.Mvc.ViewFeatures;
#nullable restore
#line 1 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\_ViewImports.cshtml"
using DiligenceReportCreation;

#line default
#line hidden
#nullable disable
#nullable restore
#line 2 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\_ViewImports.cshtml"
using DiligenceReportCreation.Models;

#line default
#line hidden
#nullable disable
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"53914d767f02b7acc15b4d5f9efb12ff5af27867", @"/Views/Diligence/GenerateFile.cshtml")]
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"48858e2aaf9df0ce88353aa7ba0222a920296d84", @"/Views/_ViewImports.cshtml")]
    public class Views_Diligence_GenerateFile : global::Microsoft.AspNetCore.Mvc.Razor.RazorPage<dynamic>
    {
        #pragma warning disable 1998
        public async override global::System.Threading.Tasks.Task ExecuteAsync()
        {
            WriteLiteral("\r\n");
#nullable restore
#line 2 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\GenerateFile.cshtml"
  
    ViewData["Title"] = "GenerateFile";

#line default
#line hidden
#nullable disable
            WriteLiteral("\r\n<h1>GenerateFile</h1>\r\n\r\n<p>\r\n    File is Ready to Download <button");
            BeginWriteAttribute("onclick", " onclick=\"", 119, "\"", 204, 3);
            WriteAttributeValue("", 129, "location.href=\'", 129, 15, true);
#nullable restore
#line 9 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\GenerateFile.cshtml"
WriteAttributeValue("", 144, Url.Action("DownloadFile", "Diligence",new { id="myLink"}), 144, 59, false);

#line default
#line hidden
#nullable disable
            WriteAttributeValue("", 203, "\'", 203, 1, true);
            EndWriteAttribute();
            WriteLiteral(@">Download File</button>
</p>
<script>
$(""#myLink"").click(function(e){

    e.preventDefault();
    $.ajax({

        url:$(this).attr(""href""), // comma here instead of semicolon
        success: function(){
            alert(""File downloaded"");  // or any other indication if you want to show            
        }

    });

});
</script>");
        }
        #pragma warning restore 1998
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.ViewFeatures.IModelExpressionProvider ModelExpressionProvider { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.IUrlHelper Url { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.IViewComponentHelper Component { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.Rendering.IJsonHelper Json { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.Rendering.IHtmlHelper<dynamic> Html { get; private set; }
    }
}
#pragma warning restore 1591
