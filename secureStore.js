import keytar from "keytar";

const SERVICE = "CircanaDashboard";

export async function saveCredential(key, value) {
  await keytar.setPassword(SERVICE, key, value);
}

export async function getCredential(key) {
  return await keytar.getPassword(SERVICE, key);
}

export async function deleteCredential(key) {
  await keytar.deletePassword(SERVICE, key);
}
