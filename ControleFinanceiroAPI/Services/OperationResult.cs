namespace ControleFinanceiroAPI.Services;

public sealed class OperationResult
{
    private OperationResult(bool success, int statusCode, string message, object? data)
    {
        Success = success;
        StatusCode = statusCode;
        Message = message;
        Data = data;
    }

    public bool Success { get; }
    public int StatusCode { get; }
    public string Message { get; }
    public object? Data { get; }

    public static OperationResult Ok(object data, string message = "Sucesso")
    {
        return new OperationResult(true, StatusCodes.Status200OK, message, data);
    }

    public static OperationResult BadRequest(string message)
    {
        return new OperationResult(false, StatusCodes.Status400BadRequest, message, null);
    }

    public static OperationResult Error(string message)
    {
        return new OperationResult(false, StatusCodes.Status500InternalServerError, message, null);
    }
}
