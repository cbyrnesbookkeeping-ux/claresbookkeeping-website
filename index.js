require("isomorphic-fetch");

const { Client } = require("@microsoft/microsoft-graph-client");
const { ClientSecretCredential } = require("@azure/identity");

module.exports = async function (context, req) {

    try {

        const {
            name,
            business,
            email,
            phone,
            message,
            website
        } = req.body || {};

        // Honeypot anti-spam field
        if (website) {
            context.res = {
                status: 400,
                body: {
                    success: false,
                    message: "Spam detected"
                }
            };
            return;
        }

        // Validation
        if (!name || !email || !message) {
            context.res = {
                status: 400,
                body: {
                    success: false,
                    message: "Required fields missing"
                }
            };
            return;
        }

        const credential = new ClientSecretCredential(
            process.env.GRAPH_TENANT_ID,
            process.env.GRAPH_CLIENT_ID,
            process.env.GRAPH_CLIENT_SECRET
        );

        const graphClient = Client.init({
            authProvider: {
                getAccessToken: async () => {

                    const token =
                      await credential.getToken(
                        "https://graph.microsoft.com/.default"
                      );

                    return token.token;
                }
            }
        });

        const emailBody = `
New Website Enquiry

Name: ${name}

Business: ${business || ""}

Email: ${email}

Phone: ${phone || ""}

Message:
${message}
`;

        await graphClient
            .api("/users/clare@claresbookkeeping.co.uk/sendMail")
            .post({
                message: {
                    subject: "New Website Enquiry",
                    body: {
                        contentType: "Text",
                        content: emailBody
                    },
                    toRecipients: [
                        {
                            emailAddress: {
                                address:
                                  "clare@claresbookkeeping.co.uk"
                            }
                        }
                    ]
                }
            });

        context.res = {
            status: 200,
            body: {
                success: true,
                message:
                  "Thank you. Your enquiry has been sent."
            }
        };

    } catch (error) {

        context.log.error(error);

        context.res = {
            status: 500,
            body: {
                success: false,
                message:
                  "Unable to send enquiry"
            }
        };
    }
};