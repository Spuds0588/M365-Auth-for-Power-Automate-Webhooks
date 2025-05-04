# M365-Auth-for-Power-Automate-Webhooks
SPA for M365 Authentication to be used with power automate webhooks

**Version:** 1.0

## Overview

`Auth.html` is a single-page HTML application designed to provide a secure way to trigger Power Automate (or other) HTTP webhooks that require user authentication via Microsoft 365 / Azure Active Directory.

It acts as a client-side bridge:
1.  A user navigates to `Auth.html` with parameters, including the target webhook URL.
2.  The page uses the Microsoft Authentication Library (MSAL.js 2.x) to ensure the user is authenticated against a configured Azure AD tenant.
3.  If not logged in, the user is redirected to the standard Microsoft login page.
4.  Upon successful authentication, the page acquires an M365 ID Token.
5.  It then redirects the user's browser to the target webhook URL, appending the acquired ID Token and any other specified parameters as query string arguments.

This allows Power Automate flows (or other services) triggered by "When an HTTP request is received" to verify the identity of the user initiating the request by validating the received ID Token.

## Features

*   **Secure Triggering:** Ensures only authenticated M365 users from your tenant can trigger the webhook.
*   **Simple Deployment:** Single HTML file with no backend required. Host it anywhere (Azure Static Web Apps, Blob Storage, GitHub Pages, internal server).
*   **Client-Side Only:** All logic runs in the user's browser.
*   **Configurable:** Easily set Azure AD App details, parameter names, and error messages via JavaScript variables.
*   **Standard M365 Login:** Uses MSAL.js for a familiar and robust authentication experience.
*   **Passes ID Token:** Forwards the user's M365 ID Token to the target webhook for validation.
*   **Parameter Forwarding:** Forwards other URL parameters passed to `Auth.html`.
*   **Minimal Dependencies:** Relies only on vanilla HTML/CSS/JS and MSAL.js loaded via CDN.
*   **User Feedback:** Provides loading indicators and clear error messages.
*   **Open Source:** Intended as a template for easy adoption (MIT License).

## How it Works (User Flow)

1.  **Navigate:** User accesses the `Auth.html` URL, like:
    `https://your-hosting.com/Auth.html?webhookUrl=ENCODED_POWER_AUTOMATE_URL¶m1=value1`
2.  **Check Auth:** The page checks for an active M365 session using MSAL.js.
3.  **Login (If Needed):** If no session exists, the user is redirected to Microsoft to log in. After login, they are redirected back to `Auth.html`.
4.  **Get Token:** MSAL.js acquires the user's ID Token.
5.  **Show Loading:** A brief "Authenticating..." message appears.
6.  **Construct URL:** The script builds the final webhook URL:
    `DECODED_POWER_AUTOMATE_URL?id_token=JWT_TOKEN_STRING¶m1=value1`
7.  **Redirect:** The user's browser is redirected to this final URL (`window.location.href`).
8.  **Webhook Execution:** The browser sends a GET request to the Power Automate webhook. The Power Automate flow receives the request, including the ID token and other parameters.
9.  **Response:** The Power Automate flow processes the request (including validating the token) and sends back an HTTP response, which the user's browser then renders.

## Prerequisites

Before using `Auth.html`, you need:

1.  **Azure Subscription:** To register an application in Azure Active Directory.
2.  **Azure AD App Registration:**
    *   Create an App Registration in your Azure AD tenant.
    *   Configure it for **Single-page application (SPA)** platform.
    *   Set the **Redirect URI** to the exact URL where you will host `Auth.html` (e.g., `https://your-hosting.com/Auth.html`).
    *   Note down the **Application (client) ID** and **Directory (tenant) ID**.
3.  **Power Automate Flow:**
    *   A Power Automate flow using the **"When an HTTP request is received"** trigger (requires a Premium license).
    *   Method must be **GET**.
    *   Note down the **HTTP POST URL** generated when you save the flow. This is your target webhook URL.
    *   **Crucially:** The flow *must* include steps to **validate the incoming ID token** (see Power Automate Setup section below).
4.  **Web Hosting:** A place to host the `Auth.html` file accessible via HTTP or HTTPS (e.g., Azure Static Web Apps, Azure Blob Storage static website, GitHub Pages, internal web server). *It cannot be run directly from the local filesystem (`file:///...`) due to MSAL limitations.*

## Setup (`Auth.html` Configuration)

1.  Download the `Auth.html` file.
2.  Open `Auth.html` in a text editor.
3.  Locate the `<!-- CONFIGURATION VARIABLES -->` section within the `<script>` block.
4.  Modify the following JavaScript variables:
    *   `msalClientId`: **Required.** Set this to your Azure AD App Registration's **Application (client) ID**.
    *   `msalTenantId`: **Required.** Set this to your Azure AD **Directory (tenant) ID**.
    *   `msalRedirectUrl`: **Required.** Set this to the **exact URL** where `Auth.html` will be hosted. This *must* match the Redirect URI configured in your Azure AD App Registration. `window.location.origin + window.location.pathname` often works if the file is at the root or a simple path.
    *   `webhookUrlParamName`: (Default: `"webhookUrl"`) The name of the URL parameter you will use to pass the *encoded* Power Automate webhook URL to `Auth.html`.
    *   `idTokenParamName`: (Default: `"id_token"`) The name of the URL parameter that will be appended to the target webhook URL, containing the user's M365 ID token.
    *   `authErrorTitle`: (Default: `"Authentication Failed"`) The title displayed if login fails.
    *   `authErrorMessage`: (Default: `"Login was unsuccessful..."`) The detailed message displayed if login fails. Customize with your support contact info.
    *   `configErrorTitle`: (Default: `"Configuration Error"`) Title for errors related to missing URL parameters.
    *   `configErrorMessage`: (Default: `"...parameter is missing..."`) Message displayed if the `webhookUrlParamName` is not found in the initial URL.
5.  Save the changes.

## Deployment

1.  Upload the modified `Auth.html` file to your chosen web hosting location.
2.  Ensure the hosting URL **exactly** matches the `msalRedirectUrl` variable in the file *and* the Redirect URI configured in your Azure AD App Registration.

    *Example Hosting Options:*
    *   **Azure Static Web Apps:** Ideal for simple, scalable hosting.
    *   **Azure Blob Storage (Static Website):** Cost-effective for static files.
    *   **GitHub Pages:** Free for public repositories.
    *   **Traditional Web Server:** IIS, Apache, Nginx, etc.
    *   **Local Testing:** Use a simple local server like `python -m http.server 8080` (access via `http://localhost:8080/Auth.html`) - ensure the port and path match your Azure AD config for testing.

## Usage

1.  **Get your Power Automate Webhook URL:** From the "When an HTTP request is received" trigger in your Power Automate flow.
2.  **URL-Encode the Webhook URL:** This is crucial because the webhook URL itself contains special characters. You can use an online tool or JavaScript's `encodeURIComponent()` function:
    ```javascript
    // In browser console:
    encodeURIComponent("YOUR_POWER_AUTOMATE_WEBHOOK_URL")
    ```
    Copy the resulting encoded string.
3.  **Construct the Link:** Create the URL that users will click:
    ```
    https://[Your_Auth.html_Hosting_URL]/Auth.html?[webhookUrlParamName]=[ENCODED_POWER_AUTOMATE_URL]&[otherParamName]=[otherValue]
    ```
    Replace:
    *   `[Your_Auth.html_Hosting_URL]` with the actual URL where you deployed `Auth.html`.
    *   `[webhookUrlParamName]` with the value of the `webhookUrlParamName` variable you configured (default is `webhookUrl`).
    *   `[ENCODED_POWER_AUTOMATE_URL]` with the URL-encoded string from step 2.
    *   `&[otherParamName]=[otherValue]` (Optional): Add any other query parameters your Power Automate flow needs. They will be passed along.

    **Example:**
    ```
    https://auth.mycompany.com/Auth.html?webhookUrl=https%3A%2F%2Fprod-....logic.azure.com%3A443%2Fworkflow....6sv%3D1.0%26sig%3D....&requestId=12345
    ```
4.  **Share the Link:** Provide this constructed link to your users. When they click it, the authentication and redirect process will begin.

## Power Automate Flow Setup (Crucial Security)

Your Power Automate flow triggered by "When an HTTP request is received" **MUST** validate the incoming ID token to ensure the request is legitimate. **Do not skip this step.**

1.  **Trigger:** "When an HTTP request is received" (Method: GET).
2.  **Get Parameters:** Access the passed parameters using expressions:
    *   ID Token: `triggerOutputs()['queries'][variables('idTokenParamName')]` (Replace `variables('idTokenParamName')` with the actual name, e.g., `'id_token'`).
    *   Other Params: `triggerOutputs()['queries']['yourOtherParamName']`
3.  **Validate ID Token:**
    *   **Parse the Token:** Use a "Parse JSON" action. The Schema can often be generated from a sample payload (copy a sample ID token received during testing). *Note: JWTs are Base64Url encoded segments separated by dots. You might only need to parse the payload (the middle segment) after decoding, or use specific JWT parsing actions if available.* A simpler approach for basic checks without full cryptographic validation (less secure, use with caution):
        *   Use `split()` on the token string by `.` to get the payload (second element).
        *   Use `base64ToString()` on the payload segment.
        *   Use `json()` to parse the resulting string.
    *   **Check Claims (Minimum):** Add "Condition" actions or expressions to verify:
        *   **Issuer (`iss`):** Should match `https://login.microsoftonline.com/[Your_Tenant_ID]/v2.0`.
            *   Expression: `equals(outputs('Parse_JSON_Payload')?['body']?['iss'], 'https://login.microsoftonline.com/YOUR_TENANT_ID/v2.0')`
        *   **Audience (`aud`):** Should match your `Auth.html` Azure AD App's **Application (client) ID**.
            *   Expression: `equals(outputs('Parse_JSON_Payload')?['body']?['aud'], 'YOUR_CLIENT_ID')`
        *   **Tenant ID (`tid`):** Should match your Azure AD **Directory (tenant) ID**.
            *   Expression: `equals(outputs('Parse_JSON_Payload')?['body']?['tid'], 'YOUR_TENANT_ID')`
        *   **Expiration (`exp`):** Should be in the future. Compare the `exp` (Unix timestamp) with the current time (`utcNow('U')`).
            *   Expression: `greater(outputs('Parse_JSON_Payload')?['body']?['exp'], utcNow('U'))`
        *   **(Optional) User (`oid` or `upn`):** Check if the Object ID or User Principal Name matches an allowed user/group if needed.
    *   **If Validation Fails:** Terminate the flow with an unauthorized status.
4.  **Add Your Flow Logic:** Perform the actions the flow is intended for.
5.  **Response Action:** Add a "Response" action at the end to send feedback to the user's browser (e.g., Status Code 200, a simple HTML confirmation message in the Body).

**Note:** For robust validation, consider using Azure Logic Apps' JWT validation actions or calling an Azure Function that uses proper libraries (like Microsoft.IdentityModel.Tokens) to cryptographically verify the token signature. The basic claim checks above are a starting point but don't verify the signature.

## Security Considerations

*   **Client-Side:** All `Auth.html` logic is visible in the browser. No secrets should ever be placed here.
*   **ID Token Validation:** The security of this solution *depends entirely* on the Power Automate flow correctly and strictly validating the received `id_token` claims (`iss`, `aud`, `tid`, `exp`, signature). **Failure to validate the token renders the authentication pointless.**
*   **HTTPS:** Always host `Auth.html` and your webhook endpoint over HTTPS.
*   **Azure AD App:** Configure the Azure AD App Registration with the least privilege necessary.

## Troubleshooting

*   **"Configuration Error" page:** Check that the URL you used includes the correct parameter name (default `webhookUrl`) and that its value (the encoded Power Automate URL) is present and correctly encoded.
*   **"Authentication Failed" page:**
    *   Check the browser's developer console (F12) for specific MSAL errors.
    *   Ensure `msalRedirectUrl` in `Auth.html` exactly matches the Redirect URI in the Azure AD App Registration.
    *   Verify `msalClientId` and `msalTenantId` are correct.
    *   Ensure the user is logging in with an account from the correct tenant (`msalTenantId`).
    *   Check if Conditional Access policies are interfering.
*   **Redirect Loops / Errors after Login:** Almost always an issue with `msalRedirectUrl` not matching the Azure AD configuration *exactly*. Check protocols (HTTP vs HTTPS), trailing slashes, path casing.
*   **Power Automate Not Triggering:**
    *   Verify the `webhookUrl` parameter value was correctly URL-encoded.
    *   Check the Power Automate flow's run history for trigger errors.
    *   Ensure the constructed URL in the browser's address bar (after redirection) looks correct.
*   **Power Automate Flow Fails:**
    *   Check the flow's run history for detailed errors.
    *   Verify the token validation logic and expressions.
    *   Check if all expected parameters (`id_token`, other custom params) are being received correctly.

## License

This project is licensed under the MIT License - see the LICENSE file for details (or add MIT License text here).
