namespace BlazorUI.Services;

public class AuthState
{
    public bool IsLoggedIn { get; private set; }

    public string? UserName { get; private set; }

    public void SignIn(string username)
    {
        UserName = username;
        IsLoggedIn = true;
    }

    public void SignOut()
    {
        UserName = null;
        IsLoggedIn = false;
    }
}
