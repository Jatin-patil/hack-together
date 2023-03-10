using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.SendMail;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.TenantId, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }

    public static Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me.GetAsync(u =>
            {
                // Only request specific properties
                u.QueryParameters.Select = new String[] { "DisplayName", "Mail", "UserPrincipalName" };
            });
    }

    public async static Task<string?> CreateTeamsMeet()
    {
        _ = _userClient ??
           throw new System.NullReferenceException("Graph has not been initialized for user auth");

        Console.WriteLine("What is the Subject of your meeting?");
        var sub = Console.ReadLine();

        Console.WriteLine("What should be the start datetime(yyyy-mm-dd HH:mm:ss) for the meeting?");
        var srtTime = Console.ReadLine();

        Console.WriteLine("For how many hours?");
        var hrs = Console.ReadLine();

        var requestBody = new OnlineMeeting
        {
            StartDateTime = DateTime.Parse(srtTime),
            EndDateTime = DateTime.Parse(srtTime).AddHours(Int32.Parse(hrs)),
            Subject = sub
        };
        var result = await _userClient.Me.OnlineMeetings.PostAsync(requestBody);
        return result?.JoinWebUrl;
    }
}

