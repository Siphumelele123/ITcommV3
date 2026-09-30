namespace BlazorUI.Services;

public enum LoginOutcome
{
    Success,
    IncorrectPassword, // mirrors: IsvalidUserPass() fails, IsvalidUser() succeeds
    IncorrectUser,     // mirrors: IsvalidUser() also fails
    InvalidOtp         // OTP field wasn't a whole number
}

public interface IAuthService
{
    Task<LoginOutcome> LoginAsync(
        string username,
        string password,
        string otp);
}