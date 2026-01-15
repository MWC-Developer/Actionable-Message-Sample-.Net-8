# Actionable Message Sample

ActionableMessageSample is a .NET 8 web API that validates and responds to Outlook Actionable Messages. It demonstrates how to verify actionable message tokens, process user actions, and send adaptive card responses, serving as a starting point for integrating Microsoft 365 Actionable Messages into existing workflows.

## Prerequisites
- .NET 8 SDK
- Access to Microsoft Entra ID with permissions to create app registrations
- Access to the [Actionable Email Developer Dashboard](https://learn.microsoft.com/en-us/outlook/actionable-messages/email-dev-dashboard)
- SSL certificate that matches the public endpoint used by the application

## Setup
1. **Create an Entra ID app registration**
   - Follow the steps in the [Enable Entra token for Actionable Messages guide](https://learn.microsoft.com/en-us/outlook/actionable-messages/enable-entra-token-for-actionable-messages).
   - Note the `Application (client) ID`, `Directory (tenant) ID`, and configure a client secret.
   - Add the required API permissions for Outlook Actionable Messages and grant admin consent.
2. **Configure the Actionable Email Developer Dashboard**
   - Sign in to the dashboard and create an Actionable Message provider entry using the same app registration and public HTTPS endpoint.
   - Obtain approval for the Actionable Message (granted by an Exchange Online administrator except for global registrations).
3. **Update application settings**
   - Populate `ActionableMessageSender/appsettings.local.json` (or the relevant environment file) with the identifiers from the Entra ID registration (`OriginatorId`, `EntraTenantId`, `EntraClientId`, `EntraAudience`, `EntraAuthorityHost`).
   - Configure the HTTPS endpoint settings under the `Kestrel` section and ensure the certificate is installed in the specified store.
4. **Run the application**
   - Restore dependencies and start the API: `dotnet run --project ActionableMessageSender`.
   - Expose the HTTPS endpoint publicly (for example via reverse proxy or Azure App Service) so Outlook can reach it.
5. **Test the flow**
   - Use the included console sender project to generate an actionable message, ensuring it targets the HTTPS endpoint you exposed.
   - Open the message in an Outlook client (desktop or web) and trigger the actionable card to verify tokens, callbacks, and response payloads.

Following these steps completes the Entra ID and Actionable Email Dashboard configuration required for the sample to run end-to-end.
