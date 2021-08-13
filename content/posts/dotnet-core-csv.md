---
title: ".Net Core csv"
date: 2021-08-13T09:17:34+08:00
draft: true
---
## .Net Core 產生csv檔 

1. 自訂IActionResult Utf8ForExcelCsvResult
```csharp
public class Utf8ForExcelCsvResult : IActionResult
{
    public string Content { get; set; }
    public string ContentType { get; set; }
    public string FileName { get; set; }

    public Task ExecuteResultAsync(ActionContext context)
    {
        var Response = context.HttpContext.Response;
        Response.Headers["Content-Type"] = this.ContentType;
        Response.Headers["Content-Disposition"] = $"attachment; filename={this.FileName}; filename*=UTF-8''{this.FileName}";
        using (var sw = new StreamWriter(Response.Body, Encoding.UTF8))
        {
            sw.Write(Content);
        }
        return Task.CompletedTask;
    }
}
```
&nbsp;

2. 產生csv檔的Contorller
```csharp
public IActionResult GetCSV()
{
    var weatherForecasts = Get();
    var builder = new StringBuilder();
    builder.AppendLine("日期,攝氏,華氏,summary");
    foreach (var item in weatherForecasts)
    {
        builder.AppendLine($"{item.Date},{item.TemperatureC},{item.TemperatureF},{item.Summary}");
    }

    return new Utf8ForExcelCsvResult()
    {
        Content = builder.ToString(),
        ContentType = "text/csv;",
        FileName = "WeatherForecast.csv",
    };
}
```
&nbsp;

問題: 從 ASP.NET Core 3.0 開始，預設會停用同步伺服器作業。
&nbsp;

https://docs.microsoft.com/zh-tw/dotnet/core/compatibility/aspnetcore#http-synchronous-io-disabled-in-all-servers
```
System.InvalidOperationException
  HResult=0x80131509
  Message=Synchronous operations are disallowed. Call WriteAsync or set AllowSynchronousIO to true instead.
  Source=Microsoft.AspNetCore.Server.IIS
  StackTrace: 
   at Microsoft.AspNetCore.Server.IIS.Core.HttpResponseStream.Write(Byte[] buffer, Int32 offset, Int32 count)
   at Microsoft.AspNetCore.Server.IIS.Core.WrappingStream.Write(Byte[] buffer, Int32 offset, Int32 count)
   at System.IO.Stream.Write(ReadOnlySpan`1 buffer)
   at System.IO.StreamWriter.Flush(Boolean flushStream, Boolean flushEncoder)
   at System.IO.StreamWriter.Dispose(Boolean disposing)
   at System.IO.TextWriter.Dispose()
   at DemoExcel.Controllers.Utf8ForExcelCsvResult.ExecuteResultAsync(ActionContext context) in D:\vs2019\Lab\DemoExcel\Controllers\WeatherForecastController.cs:line 85
   at Microsoft.AspNetCore.Mvc.Infrastructure.ResourceInvoker.InvokeResultAsync(IActionResult result)
   at Microsoft.AspNetCore.Mvc.Infrastructure.ResourceInvoker.ResultNext[TFilter,TFilterAsync](State& next, Scope& scope, Object& state, Boolean& isCompleted)
   at Microsoft.AspNetCore.Mvc.Infrastructure.ResourceInvoker.InvokeNextResultFilterAsync[TFilter,TFilterAsync]()
```
解法1: 在starup增加
```csharp
public void ConfigureServices(IServiceCollection services)
{
    // If using Kestrel:
    services.Configure<KestrelServerOptions>(options =>
    {
        options.AllowSynchronousIO = true;
    });

    // If using IIS:
    services.Configure<IISServerOptions>(options =>
    {
        options.AllowSynchronousIO = true;
    });
}
```
解法2: 修改Utf8ForExcelCsvResult的ExecuteResultAsync方法為非同步
```csharp
public async Task ExecuteResultAsync(ActionContext context)
{
    var Response = context.HttpContext.Response;
    Response.Headers["Content-Type"] = this.ContentType;
    Response.Headers["Content-Disposition"] = $"attachment; filename={this.FileName}; filename*=UTF-8''{this.FileName}";

    // 讓編碼格式變成utf-8 whit BOM
    await context.HttpContext.Response.WriteAsync("\uFEFF");
    await context.HttpContext.Response.WriteAsync(Content, Encoding.UTF8);
}
```
&nbsp;

參考:
&nbsp;

+ https://stackoverflow.com/questions/52491983/excel-csv-encoding-issues
+ https://stackoverflow.com/questions/59388629/what-is-the-usage-of-executeresultasyncactioncontext-method
+ https://www.red-gate.com/simple-talk/development/dotnet-development/action-control-asp-net-core/
+ https://stackoverflow.com/questions/47735133/asp-net-core-synchronous-operations-are-disallowed-call-writeasync-or-set-all
+ https://stackoverflow.com/questions/1005530/how-to-add-encoding-information-to-the-response-stream-in-asp-net
