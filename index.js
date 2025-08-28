import { PublicClientApplication } from "@azure/msal-node";
import fetch from "node-fetch";

// Configuración de tu app en Azure
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
    //  Pedir URL de login
    const authCodeUrlParameters = {
      scopes: ["User.Read", "Contacts.ReadWrite"], 
      redirectUri: "http://localhost:3000"
    };

    const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
    console.log(" Abre este enlace en el navegador y copia el código de autorización:");
    console.log(authUrl);

    // Pausa manual: pega aquí el "authorization code" que obtienes del navegador
    const readline = await import("readline");
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

    rl.question("Pega aquí el código de autorización: ", async (authCode) => {
      // Intercambiar el código por un token
      const tokenResponse = await pca.acquireTokenByCode({
        code: authCode,
        scopes: ["User.Read", "Contacts.ReadWrite"],
        redirectUri: "http://localhost:3000"
      });

      const accessToken = tokenResponse.accessToken;
      console.log("Token obtenido!");

      //Crear un contacto en Outlook con categoría
      const newContact = {
        givenName: "Cristian",
        surname: "Estrada",
        emailAddresses: [
          {
            address: "Cristian.Estrada@ejemplo.com",
            name: "Cristian Estrada"
          }
        ],
        businessPhones: ["+57 3000000000"],
        categories: ["Amigos"] //Categoría que le asignamos
      };

      const response = await fetch("https://graph.microsoft.com/v1.0/me/contacts", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify(newContact)
      });

      if (!response.ok) {
        console.error(" Error al crear contacto:", await response.text());
      } else {
        console.log("Contacto creado con categoría!");
      }

      rl.close();
    });

  } catch (err) {
    console.error("Error general:", err);
  }
}

main();
