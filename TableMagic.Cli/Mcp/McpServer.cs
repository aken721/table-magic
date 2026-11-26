using System.Text.Json;
using System.Text.Json.Serialization;
using TableMagic.Cli.Skills;

namespace TableMagic.Cli.Mcp;

public class McpServer
{
    private readonly StdioTransport _transport;
    private readonly SkillManager _skillManager;

    private static readonly JsonSerializerOptions CachedJsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public McpServer(SkillManager skillManager)
    {
        _transport = new StdioTransport();
        _skillManager = skillManager;
    }

    public async Task RunAsync()
    {
        while (true)
        {
            var message = await _transport.ReadMessageAsync();
            if (message == null) break;

            JsonRpcRequest? request;
            try
            {
                request = JsonSerializer.Deserialize<JsonRpcRequest>(message, CachedJsonOptions);
            }
            catch
            {
                await _transport.SendErrorAsync(null, JsonRpcError.ParseError);
                continue;
            }

            if (request == null || string.IsNullOrEmpty(request.Method))
            {
                await _transport.SendErrorAsync(request?.Id, JsonRpcError.InvalidRequest);
                continue;
            }

            await HandleRequestAsync(request);
        }
    }

    private async Task HandleRequestAsync(JsonRpcRequest request)
    {
        try
        {
            switch (request.Method)
            {
                case "initialize":
                    await HandleInitializeAsync(request);
                    break;
                case "notifications/initialized":
                    break;
                case "tools/list":
                    await HandleToolsListAsync(request);
                    break;
                case "tools/call":
                    await HandleToolCallAsync(request);
                    break;
                default:
                    await _transport.SendErrorAsync(request.Id, JsonRpcError.MethodNotFound);
                    break;
            }
        }
        catch (Exception ex)
        {
            await _transport.SendErrorAsync(request.Id, new JsonRpcError
            {
                Code = -32603,
                Message = ex.Message
            });
        }
    }

    private async Task HandleInitializeAsync(JsonRpcRequest request)
    {

        var result = new InitializeResult
        {
            ProtocolVersion = "2024-11-05",
            Capabilities = new ServerCapabilities
            {
                Tools = new ToolCapabilities { ListChanged = false }
            },
            ServerInfo = new ServerInfo
            {
                Name = "tablemagic",
                Version = typeof(McpServer).Assembly.GetName().Version?.ToString(3) ?? "2.5.2.1"
            }
        };

        var response = new JsonRpcResponse
        {
            Id = request.Id,
            Result = result
        };
        await _transport.SendMessageAsync(response);
    }

    private async Task HandleToolsListAsync(JsonRpcRequest request)
    {
        var skills = _skillManager.GetAllTools();
        var tools = skills.Select(t => new McpTool
        {
            Name = t.Name,
            Description = t.Description,
            InputSchema = t.Parameters
        }).ToList();

        var result = new ToolsListResult { Tools = tools };
        var response = new JsonRpcResponse
        {
            Id = request.Id,
            Result = result
        };
        await _transport.SendMessageAsync(response);
    }

    private async Task HandleToolCallAsync(JsonRpcRequest request)
    {
        var callParams = request.Params is JsonElement paramsElement
            ? JsonSerializer.Deserialize<ToolCallParams>(paramsElement.GetRawText(), CachedJsonOptions)
            : null;

        if (callParams == null || string.IsNullOrEmpty(callParams.Name))
        {
            await _transport.SendErrorAsync(request.Id, JsonRpcError.InvalidParams);
            return;
        }

        var arguments = callParams.Arguments ?? new Dictionary<string, object>();
        var result = await _skillManager.ExecuteToolAsync(callParams.Name, arguments);

        var callResult = new ToolCallResult
        {
            IsError = !result.Success,
            Content = new List<ToolCallContent>()
        };

        if (result.MissingRequiredParams && result.MissingParams.Count > 0)
        {
            callResult.Content.Add(new ToolCallContent
            {
                Type = "text",
                Text = result.Error ?? ""
            });

            foreach (var mp in result.MissingParams)
            {
                callResult.Content.Add(new ToolCallContent
                {
                    Type = "text",
                    Text = $"[MISSING_PARAM] {mp.Name} ({mp.Type}): {mp.PromptHint}"
                });
            }
        }
        else
        {
            callResult.Content.Add(new ToolCallContent
            {
                Type = "text",
                Text = result.Success ? result.Content ?? "" : result.Error ?? "Unknown error"
            });
        }

        var response = new JsonRpcResponse
        {
            Id = request.Id,
            Result = callResult
        };
        await _transport.SendMessageAsync(response);
    }
}
