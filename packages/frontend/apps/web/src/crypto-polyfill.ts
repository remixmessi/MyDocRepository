/**
 * `crypto.randomUUID()` is unavailable in some browsers when the app is
 * opened over plain HTTP from a LAN IP. Keep the existing call sites working
 * while still using the browser's cryptographically secure random source when
 * it is available.
 */
const webCrypto = globalThis.crypto ?? ({} as Crypto);

if (!globalThis.crypto) {
  Object.defineProperty(globalThis, 'crypto', { configurable: true, value: webCrypto });
}

if (webCrypto && typeof webCrypto.randomUUID !== 'function') {
  Object.defineProperty(webCrypto, 'randomUUID', {
    configurable: true,
    value: () => {
      const bytes = new Uint8Array(16);
      if (typeof webCrypto.getRandomValues === 'function') {
        webCrypto.getRandomValues(bytes);
      } else {
        for (let index = 0; index < bytes.length; index += 1) bytes[index] = Math.floor(Math.random() * 256);
      }
      bytes[6] = (bytes[6] & 0x0f) | 0x40;
      bytes[8] = (bytes[8] & 0x3f) | 0x80;
      const hex = [...bytes].map(byte => byte.toString(16).padStart(2, '0')).join('');
      return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
    },
  });
}
