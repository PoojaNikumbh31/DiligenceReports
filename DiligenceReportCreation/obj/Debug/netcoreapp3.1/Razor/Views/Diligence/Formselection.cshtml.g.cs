#pragma checksum "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\Formselection.cshtml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "072e635f8b8053362895d7f0a55f6b1c1ea7badc"
// <auto-generated/>
#pragma warning disable 1591
[assembly: global::Microsoft.AspNetCore.Razor.Hosting.RazorCompiledItemAttribute(typeof(AspNetCore.Views_Diligence_Formselection), @"mvc.1.0.view", @"/Views/Diligence/Formselection.cshtml")]
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
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"072e635f8b8053362895d7f0a55f6b1c1ea7badc", @"/Views/Diligence/Formselection.cshtml")]
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"48858e2aaf9df0ce88353aa7ba0222a920296d84", @"/Views/_ViewImports.cshtml")]
    public class Views_Diligence_Formselection : global::Microsoft.AspNetCore.Mvc.Razor.RazorPage<DiligenceReportCreation.Models.ReportModel>
    {
        #pragma warning disable 1998
        public async override global::System.Threading.Tasks.Task ExecuteAsync()
        {
#nullable restore
#line 2 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\Formselection.cshtml"
  
    ViewData["Title"] = "Formselection";

#line default
#line hidden
#nullable disable
            WriteLiteral("\r\n\r\n    <div class=\"text-center\">\r\n        <h1>Diligence Reports</h1>\r\n");
#nullable restore
#line 9 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\Formselection.cshtml"
         using (Html.BeginForm())
        {
            

#line default
#line hidden
#nullable disable
#nullable restore
#line 11 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\Formselection.cshtml"
       Write(Html.AntiForgeryToken());

#line default
#line hidden
#nullable disable
            WriteLiteral(@"            <div class=""form-group"">
                <br />
                <br />
                <label>Case Number</label>
                <input type=""text"" name=""CaseNumber"" required=""required"" />
                <br />
                <label>Last Name/Company Name</label>
                <input type=""text"" name=""LastName"" required=""required"" />
                <br />
");
            WriteLiteral("                <br />\r\n\r\n                <input id=\"submit\" type=\"submit\" name=\"Submit\" value=\"Search\" />\r\n            </div>\r\n");
#nullable restore
#line 27 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\Formselection.cshtml"
        }

#line default
#line hidden
#nullable disable
#nullable restore
#line 28 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\Formselection.cshtml"
         if (!(@TempData["message"] is null))
        {

#line default
#line hidden
#nullable disable
            WriteLiteral("            <p class=\"alert-danger\">");
#nullable restore
#line 30 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\Formselection.cshtml"
                               Write(TempData["message"].ToString());

#line default
#line hidden
#nullable disable
            WriteLiteral("</p>\r\n");
#nullable restore
#line 31 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Diligence\Formselection.cshtml"
        }

#line default
#line hidden
#nullable disable
            WriteLiteral("    </div>");
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
        public global::Microsoft.AspNetCore.Mvc.Rendering.IHtmlHelper<DiligenceReportCreation.Models.ReportModel> Html { get; private set; }
    }
}
#pragma warning restore 1591