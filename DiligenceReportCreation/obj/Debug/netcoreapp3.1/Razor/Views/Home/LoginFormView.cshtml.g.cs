#pragma checksum "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "3a62277bedb23f34e9e1e612da179476a1cdb9b1"
// <auto-generated/>
#pragma warning disable 1591
[assembly: global::Microsoft.AspNetCore.Razor.Hosting.RazorCompiledItemAttribute(typeof(AspNetCore.Views_Home_LoginFormView), @"mvc.1.0.view", @"/Views/Home/LoginFormView.cshtml")]
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
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"3a62277bedb23f34e9e1e612da179476a1cdb9b1", @"/Views/Home/LoginFormView.cshtml")]
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"48858e2aaf9df0ce88353aa7ba0222a920296d84", @"/Views/_ViewImports.cshtml")]
    public class Views_Home_LoginFormView : global::Microsoft.AspNetCore.Mvc.Razor.RazorPage<DiligenceReportCreation.Models.LoginViewModel>
    {
        #pragma warning disable 1998
        public async override global::System.Threading.Tasks.Task ExecuteAsync()
        {
#nullable restore
#line 2 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml"
  
    ViewData["Title"] = "LoginFormView";

#line default
#line hidden
#nullable disable
            WriteLiteral("\n    <div class=\"text-center\">\n        <h1>Login</h1>\n");
#nullable restore
#line 8 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml"
         using (Html.BeginForm("LoginAuth", "Home", FormMethod.Post))
        {

#line default
#line hidden
#nullable disable
            WriteLiteral("            <div class=\"form-group\">\n                <label>Username:</label>\n                ");
#nullable restore
#line 12 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml"
           Write(Html.TextBoxFor(LoginModel => LoginModel.Username));

#line default
#line hidden
#nullable disable
            WriteLiteral("\n            </div>\n            <div class=\"form-group\">\n                <label>Password :</label>\n                ");
#nullable restore
#line 16 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml"
           Write(Html.PasswordFor(LoginModel => LoginModel.Password));

#line default
#line hidden
#nullable disable
            WriteLiteral("\n            </div>\n");
            WriteLiteral("            <br />\n            <input id=\"Submit\" type=\"submit\" value=\"Login\" />     \n");
#nullable restore
#line 21 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml"

        }

#line default
#line hidden
#nullable disable
#nullable restore
#line 23 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml"
         if (!(@TempData["message"] is null))
        {

#line default
#line hidden
#nullable disable
            WriteLiteral("            <p class=\"alert-danger\">");
#nullable restore
#line 25 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml"
                               Write(TempData["message"].ToString());

#line default
#line hidden
#nullable disable
            WriteLiteral("</p>\n");
#nullable restore
#line 26 "C:\Users\Pooja.nikumbh\source\repos\NEW_FOLDER_FOR_PROJECT\DiligenceReportCreation(IN_Citi)\DiligenceReportCreation\Views\Home\LoginFormView.cshtml"
        }

#line default
#line hidden
#nullable disable
            WriteLiteral("    </div>\n    \n\n");
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
        public global::Microsoft.AspNetCore.Mvc.Rendering.IHtmlHelper<DiligenceReportCreation.Models.LoginViewModel> Html { get; private set; }
    }
}
#pragma warning restore 1591
