namespace t;

public class UnitTests
{
    private Tests _Tests;

    [SetUp]
    public void Setup()
    {
        _Tests = new Tests();
    }

    [Test]
    public void Login_WithValidCredentials_ShouldSucceed()
    {
        // Arrange
        string username = "test.user@domain.com";
        string password = "Strong123";

        // Act
        bool result = _Tests.LoginTest(username,password);

        // Assert
        TestContext.WriteLine($"Login result for user {username} with valid credentials: {result}");
        Assert.That(result, Is.True, "Login should succeed with valid credentials");

    }
}
