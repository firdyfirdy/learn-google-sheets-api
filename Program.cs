using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System;
using System.IO;

namespace GoogleSheetsAndCsharp
{
  class Program
  {
    /* Read more about Google Sheets Api :
     * https://developers.google.com/sheets/api/guides/concepts
     */

    static int tableWidth = 150;
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string ApplicationNames = "MoneyManager";

    /* Your Spreadsheet ID. Example: https://docs.google.com/spreadsheets/d/1x1jxZdFM4jRI7KioynABmNuftr9WjVtCNUT_5IQ7vJQ */
    static readonly string SpreadsheetId = "1x1jxZdFM4jRI7KioynABmNuftr9WjVtCNUT_5IQ7vJQ";

    /* Your Sheet Name */
    static readonly string sheet = "Money";

    static SheetsService service;
    static void Main(string[] args)
    {
      GoogleCredential credential;

      /* Client Secret generated when you create credentials from here: https://console.developers.google.com/apis/credentials */
      using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
      {
        credential = GoogleCredential.FromStream(stream)
          .CreateScoped(Scopes);
      }

      service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
      {
        HttpClientInitializer = credential,
        ApplicationName = ApplicationNames
      });

      Console.Clear();
      showHeader();
      ReadEntries();
    }

    static void showHeader()
    {
      Console.WriteLine("# Money Usage List: #");
    }

    /* Function for read data from spreadsheet */
    static void ReadEntries()
    {
      var range = $"{sheet}!C5:H44";
      var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

      var response = request.Execute();
      var values = response.Values;

      if (values != null && values.Count > 0)
      {
        PrintLine();
        PrintRow("Date", "Note", "Category", "In", "Out", "Balance");
        PrintLine();
        foreach (var row in values)
        {
          string In = row[3].ToString();
          In = String.IsNullOrWhiteSpace(In) ? "-" : In;

          string Out = row[4].ToString();
          Out = String.IsNullOrWhiteSpace(Out) ? "-" : Out;


          PrintRow(row[0].ToString(), row[1].ToString(), row[2].ToString(), In, Out, row[5].ToString());
        }
        PrintLine();
      }
      else
      {
        Console.WriteLine("No Data Found");
      }
    }
    
    /* Method for generating a "-" symbol across tableWidth */
    static void PrintLine()
    {
      Console.WriteLine(new string('-', tableWidth));
    }

    /* Method for print row and insert it to "table" */
    static void PrintRow(params string[] columns)
    {
      int width = (tableWidth - columns.Length) / columns.Length;
      string row = "|";

      foreach (string column in columns)
      {
        row += AlignCentre(column, width) + "|";
      }

      Console.WriteLine(row);
    }

    static string AlignCentre(string text, int width)
    {
      text = text.Length > width ? text.Substring(0, width - 3) + "..." : text;

      if (string.IsNullOrEmpty(text))
      {
        return new string(' ', width);
      }
      else
      {
        return text.PadRight(width - (width - text.Length) / 2).PadLeft(width);
      }
    }
  }
}
