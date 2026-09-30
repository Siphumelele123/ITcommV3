using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace BlazorUI.Services;

/// <summary>
/// Rewrite of FrmLoginAtt.IsvalidUserPass() / IsvalidUser() using plain ADO.NET,
/// because the original DataClassLegalsuiteDataContext (LINQ to SQL) does not
/// run on .NET 9. Same two-step check, same three outcomes.
/// </summary>
public class AuthService : IAuthService
{
    private readonly string _connectionString;

    public AuthService(IConfiguration config)
    {
        _connectionString = config.GetConnectionString("Legalsuite")
            ?? throw new InvalidOperationException(
                "Connection string 'Legalsuite' is missing from appsettings.json.");
    }

    public async Task<LoginOutcome> LoginAsync(
        string username,
        string password,
        string otp)
    {
        // WinForms: string newotp = txtOTP.Text; int.Parse(newotp)
        // throws on bad input there too.
        if (!int.TryParse(otp, out var otpValue))
            return LoginOutcome.InvalidOtp;

        await using var conn = new SqlConnection(_connectionString);

        await conn.OpenAsync();

        // Mirrors IsvalidUserPass():
        // LoginID + Password + EmployeeTypeID + OTP all match.
        //
        // NOTE:
        // Preserves the original asymmetric case check:
        // UPPER(LoginID) is compared against the username exactly as typed.
        // Same for Password.
        const string fullCheckSql = @"
            SELECT COUNT(1)
            FROM Employees
            WHERE UPPER(LoginID) = @User
              AND UPPER(Password) = @Pass
              AND EmployeeTypeID <> 19
              AND OTP = @Otp";

        await using (var cmd = new SqlCommand(fullCheckSql, conn))
        {
            cmd.Parameters.AddWithValue("@User", username);
            cmd.Parameters.AddWithValue("@Pass", password);
            cmd.Parameters.AddWithValue("@Otp", otpValue);

            var result = await cmd.ExecuteScalarAsync();

            var matched = Convert.ToInt32(result);

            if (matched > 0)
                return LoginOutcome.Success;
        }

        // Mirrors IsvalidUser():
        // Same check, but ignores Password.
        const string userOnlySql = @"
            SELECT COUNT(1)
            FROM Employees
            WHERE UPPER(LoginID) = @User
              AND EmployeeTypeID <> 19
              AND OTP = @Otp";

        await using (var cmd = new SqlCommand(userOnlySql, conn))
        {
            cmd.Parameters.AddWithValue("@User", username);
            cmd.Parameters.AddWithValue("@Otp", otpValue);

            var result = await cmd.ExecuteScalarAsync();

            var matched = Convert.ToInt32(result);

            return matched > 0
                ? LoginOutcome.IncorrectPassword
                : LoginOutcome.IncorrectUser;
        }
    }
}