declare module 'jwks-client' {
  export interface JwksClientOptions {
    jwksUri: string;
    cache?: boolean;
    cacheMaxAge?: number;
    rateLimit?: boolean;
    jwksRequestsPerMinute?: number;
  }

  export interface SigningKey {
    getPublicKey(): string;
  }

  export interface JwksClient {
    getSigningKey(kid: string, callback: (err: Error | null, key?: SigningKey) => void): void;
  }

  function jwksClient(options: JwksClientOptions): JwksClient;
  export = jwksClient;
}