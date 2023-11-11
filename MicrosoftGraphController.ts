import MicrosoftGraphServices from './MicrosoftGraphServices'; // Adjust the import based on your setup

class MicrosoftGraphController {
    private refreshToken: string | null = null;

    public async ReadEmailHistory(): Promise<Array<Object>> {
        const graphService = new MicrosoftGraphServices();

        // Use the class's refreshToken property
        if (!this.refreshToken) {
            console.error('No refresh token found');
            console.error('Please run the MakeNewRefreshToken method first');
        }

        const accessToken = await graphService.getAccessToken(this.refreshToken);

        let allMessages = [];
        const pageSize = 2;
        let skip = 0;
        let totalFetched = 0;

        // Convert the loop to async/await pattern to avoid hitting the rate limit
        let batchCount: number = 0;
        do {
            const data = await graphService.getMessages(accessToken, pageSize, skip);
            if (data.value && data.value.length) {
                allMessages = allMessages.concat(data.value);
                batchCount = data.value.length;
                totalFetched += batchCount;
                skip += batchCount;
            } else {
                batchCount = 0;
            }
        } while (batchCount === pageSize);
        console.log('Total Messages Fetched:', totalFetched);
        return allMessages;
    }


    //
    public async MakeNewRefreshToken(): Promise<string> {
        const graphService = new MicrosoftGraphServices();
        const url = graphService.getAuthorizationUrl();
        // In a Node.js environment, you'd typically return the URL rather than directly echoing it.
        return `<a href='${url}'>Click here to login</a>`;
    }


    /// this will be the callback url handler
    public async callbackTokenHandler(code: string): Promise<void> {
        if (!code) {
            throw new Error("Error getting code");
        }
        const graphService = new MicrosoftGraphServices();
        const response = await graphService.getAuthenticationCodes(code);

        if (response.error) {
            throw new Error(response.error_description);
        }
        this.refreshToken = response.refresh_token;
    }
}
