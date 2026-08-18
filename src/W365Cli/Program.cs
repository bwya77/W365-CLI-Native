using System.Text;
using W365Cli;

try
{
    Console.OutputEncoding = Encoding.UTF8;
}
catch (IOException)
{
    // Output is redirected to a file/pipe that doesn't support changing encoding — ignore.
}

var app = new W365CliApp();
return await app.RunAsync(args);
