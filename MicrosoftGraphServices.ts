import axios from 'axios'; // or another HTTP client library


export class MicrosoftGraphServices {

    public static email: string = process.env.EMAIL || 'your-email-address';
    public static API_TYPE: string = 'MicrosoftGraphAPI';

    private static clientId: string = process.env.CLIENT_ID || 'your-client-id';
    private static tenantId: string = process.env.TENANT_ID || 'your-tenant-id';
    private static SCOPE: string = "offline_access user.read mail.read";
    private static REDIRECT_URI: string = process.env.REDIRECT_URI || "your redirect url"; // Change as needed
    private static clientSecretKey: string = process.env.CLIENT_SECRET_KEY || 'your-client-secret-key';
    private static clientSecretValue: string = process.env.CLIENT_SECRET_VALUE || 'your-client-secret-value';

    private static selectFields: string[] = ['sender', 'subject', 'body', 'toRecipients', 'bodyPreview', 'sentDateTime'];

    public getAuthorizationUrl(): string {
        try {
            const authUri = `https://login.microsoftonline.com/${MicrosoftGraphServices.tenantId}/oauth2/v2.0/authorize?client_id=${MicrosoftGraphServices.clientId}&scope=${encodeURIComponent(MicrosoftGraphServices.SCOPE)}&redirect_uri=${encodeURIComponent(MicrosoftGraphServices.REDIRECT_URI)}&response_type=code&approval_prompt=auto`;
            return authUri;
        } catch (error) {
            throw error;
        }
    }

    public async getAuthenticationCodes(code: string): Promise<any> {
        try {
            const response = await axios.post(`https://login.microsoftonline.com/${MicrosoftGraphServices.tenantId}/oauth2/v2.0/token`, {
                client_id: MicrosoftGraphServices.clientId,
                scope: 'User.Read Mail.Read',
                code,
                redirect_uri: MicrosoftGraphServices.REDIRECT_URI,
                grant_type: 'authorization_code',
                client_secret: MicrosoftGraphServices.clientSecretValue,
            });

            return response.data;
        } catch (error) {
            throw error;
        }
    }



    public async getAccessToken(refreshToken: string): Promise<string> {
        try {
            const url = `https://login.microsoftonline.com/${MicrosoftGraphServices.tenantId}/oauth2/v2.0/token`;

            const response = await axios.post(url, new URLSearchParams({
                client_id: MicrosoftGraphServices.clientId,
                scope: 'https://graph.microsoft.com/mail.read',
                client_secret: MicrosoftGraphServices.clientSecretValue,
                grant_type: 'refresh_token',
                refresh_token: refreshToken,
            }).toString(), {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
            });

            if (response.data.access_token) {
                return response.data.access_token;
            } else {
                console.error('Error in getAccessToken Method: ', response.data.error_description);
                throw new Error(response.data.error_description);
            }
        } catch (error) {
            console.error('Error in getAccessToken Method: ', error);
            throw error;
        }
    }


    public async getMessages(accessToken: string, top: number = 10, skip: number = 0): Promise<any> {
        try {
            const selectFields = MicrosoftGraphServices.selectFields;
            const queryParams = new URLSearchParams({
                '$select': selectFields.join(','),
                '$orderby': 'sentDateTime DESC',
                '$top': top.toString(),
                '$skip': skip.toString(),
                '$count': "true",
                // '$filter': "isDraft eq false",
            });

            const url = `https://graph.microsoft.com/v1.0/me/messages?${queryParams}`;

            const response = await axios.get(url, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                },
            });

            if (response.status === 200) {
                return response.data;
            } else {
                throw new Error(`Error: ${response.data.error.message}`);
            }
        } catch (error) {
            console.error('Error in getMessages Method: ', error);
            throw error;
        }
    }

}
