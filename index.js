import { PublicClientApplication } from "@azure/msal-node";
import fetch from "node-fetch";

// Configuración de la aplicación en Azure
const config = {
  auth: {
    clientId: "b538cc43-fcc5-4639-9055-1bee60c3cc8a",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost:3000",
  }
};

const pca = new PublicClientApplication(config);

async function main() {
  try {
    // Generar la URL de autorización
    const authCodeUrlParameters = {
      scopes: ["User.Read", "Contacts.ReadWrite"],
      redirectUri: "http://localhost:3000"
    };

    const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
    console.log("Abra este enlace en el navegador y copie el código de autorización:");
    console.log(authUrl);

    // Lectura manual del código de autorización
    const readline = await import("readline");
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

    rl.question("Pegue aquí el código de autorización: ", async (authCode) => {
      try {
        // Intercambiar el código por un token
        const tokenResponse = await pca.acquireTokenByCode({
          code: authCode,
          scopes: ["User.Read", "Contacts.ReadWrite"],
          redirectUri: "http://localhost:3000"
        });

        const accessToken = tokenResponse.accessToken;
        console.log("Token obtenido correctamente");

        // Lista de contactos a crear
        const newContacts = [
          {
            givenName: "Cristian",
            surname: "Estrada",
            emailAddresses: [
              {
                address: "Cristian.Estrada@ejemplo.com",
                name: "Cristian Estrada"
              }
            ],
            businessPhones: ["+57 3000000000"],
            categories: ["Amigos"]
          },
          {
            givenName: "Laura",
            surname: "García",
            emailAddresses: [
              {
                address: "Laura.Garcia@ejemplo.com",
                name: "Laura García"
              }
            ],
            businessPhones: ["+57 3011111111"],
            categories: ["Trabajo"]
          },
          {
            givenName: "Andrés",
            surname: "Pérez",
            emailAddresses: [
              {
                address: "Andres.Perez@ejemplo.com",
                name: "Andrés Pérez"
              }
            ],
            businessPhones: ["+57 3022222222"],
            categories: ["Familia"]
          }
        ];

        for (const contact of newContacts) {
          const response = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${accessToken}`,
              "Content-Type": "application/json"
            },
            body: JSON.stringify(contact)
          });

          if (!response.ok) {
            console.error("Error al crear contacto:", await response.text());
          } else {
            console.log(`Contacto ${contact.givenName} creado correctamente`);
          }
        }
      } catch (error) {
        console.error("Error al intercambiar el código:", error);
      } finally {
        rl.close();
      }
    });

  } catch (err) {
    console.error("Error general:", err);
  }
}

main()