# SPFX Webpart for Graph API Integration

Welcome to the SPFX Webpart for Graph API Integration! This solution is designed to help beginners and developers working with Microsoft Graph API. It includes three components: User, Mail, and Calendar endpoints. Below, you'll find detailed instructions on how to use each component, along with important tips and best practices.

## Components

### User Endpoint

- **Purpose**: Fetch logged-in user data and any user data by email.
- **Usage**:
  - To get the logged-in user data: `GET /me`
  - To get user data by email: `GET /users/{email}`

### Calendar Endpoint

- **Purpose**: Fetch events from Outlook and create new events.
- **Usage**:
  - To fetch events: `GET /me/events`
  - To create a new event: `POST /me/events`

### Mail Endpoint

- **Purpose**: Fetch mails from Outlook and send emails.
- **Usage**:
  - To fetch mails: `GET /me/messages`
  - To send an email: `POST /me/sendMail`

## Libraries Used

- **PnP JS**: Version 4.11.0
- **Node**: Version 18.23.0

## Reference Links

1. [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. [PnP JS](https://pnp.github.io/pnpjs/)

## Important Notes

1. **Global Admin Permission**: You must have global admin permission to approve Graph API from the admin portal.
2. **Security Consideration**: When you approve a Graph API endpoint, it is not limited to your solution only. It will be granted tenant-wide for solutions using the same API. This poses a security risk, so be cautious.
3. **Follow Official Documentation**: Always refer to the official Microsoft documentation for the most updated information instead of relying on third-party sources.

## Video Tutorial

For a detailed walkthrough, check out my YouTube tutorial. Feel free to reach out if you have any questions or need further assistance. Happy coding! 😊
