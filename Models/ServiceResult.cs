namespace EhrIdentityAudit.Models;

public class ServiceResult<T>
{
    public bool IsSuccess { get; init; }
    public T? Value { get; init; }
    public string? ErrorMessage { get; init; }

    public static ServiceResult<T> Success(T value) =>
        new() { IsSuccess = true, Value = value };

    public static ServiceResult<T> Fail(string message) =>
        new() { IsSuccess = false, ErrorMessage = message };
}