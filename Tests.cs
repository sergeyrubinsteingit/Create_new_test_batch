using NUnit.Framework;
using System.Text.RegularExpressions;

[TestFixture]
public class Tests
{
    // This is a mock "login" method you'd normally call your actual service for
    public bool LoginTest(string username, string password)
    {
        // Simplified validation:
        // username must be non-empty and contain '@'
        // password must be at least 6 chars and contain at least one number
        if (string.IsNullOrWhiteSpace(username) || !username.Contains("@"))
            return false;

        if (string.IsNullOrWhiteSpace(password) || password.Length < 6)
            return false;

        // Must contain a digit
        if (!Regex.IsMatch(password, @"\d"))
            return false;

        // Simulate successful login
        return username == "test.user@domain.com" && password == "Strong123";
    }

    [Test]
    public void Sanity()
    {
        Assert.Pass("Sanity OK");
    }

    [Test]
    public void Login_WithValidCredentials_ShouldSucceed()
    {
        // Arrange
        string username = "test.user@domain.com";
        string password = "Strong123";

        // Act
        bool result = LoginTest(username, password);

        // Assert
        TestContext.WriteLine($"Login result for user {username} with valid credentials: {result}");
        Assert.That(result, Is.True, "Login should succeed with valid credentials");

    }

    [Test]
    public void Login_WithInvalidPassword_ShouldFail()
    {
        // Arrange
        string username = "test.user@domain.com";
        string password = "weak"; // too short, no digits

        // Act
        bool result = LoginTest(username, password);

        // Assert
        TestContext.WriteLine($"Login result for invalid password {password}: {result}");
        Assert.That(result, Is.False, "Login should fail with invalid password");
    }

    [Test]
    public void Login_WithInvalidUsername_ShouldFail()
    {
        // Arrange
        string username = "invalid_user"; // missing '@'
        string password = "Valid123";

        // Act
        bool result = LoginTest(username, password);

        // Assert
        TestContext.WriteLine($"Login result for invalid username {username}: {result}");
        Assert.That(result, Is.False, "Login should fail with invalid username");
    }
}
