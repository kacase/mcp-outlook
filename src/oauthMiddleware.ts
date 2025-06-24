import { Request, Response, NextFunction } from 'express';
import jwt from 'jsonwebtoken';
import jwksClient from 'jwks-client';

interface TokenPayload {
  aud: string;
  iss: string;
  sub: string;
  scp?: string;
  roles?: string[];
  exp: number;
  iat: number;
}

export class MicrosoftTokenValidator {
  private jwksClient: any;
  private authority: string;
  private clientId: string;

  constructor() {
    this.authority = process.env.AUTHORITY || 'common';
    this.clientId = process.env.CLIENT_ID!;
    
    this.jwksClient = jwksClient({
      jwksUri: `https://login.microsoftonline.com/${this.authority}/discovery/v2.0/keys`,
      cache: true,
      rateLimit: true,
      jwksRequestsPerMinute: 10
    });
  }

  private async getSigningKey(kid: string): Promise<string> {
    return new Promise((resolve, reject) => {
      this.jwksClient.getSigningKey(kid, (err: Error | null, key: any) => {
        if (err) return reject(err);
        resolve(key!.getPublicKey());
      });
    });
  }

  async validateToken(token: string): Promise<TokenPayload> {
    return new Promise(async (resolve, reject) => {
      try {
        // Decode token header to get kid
        const decoded = jwt.decode(token, { complete: true });
        if (!decoded || typeof decoded === 'string' || !decoded.header.kid) {
          return reject(new Error('Invalid token format'));
        }

        // Get signing key
        const signingKey = await this.getSigningKey(decoded.header.kid);

        // Verify and decode token
        const payload = jwt.verify(token, signingKey, {
          audience: this.clientId,
          issuer: `https://login.microsoftonline.com/${this.authority}/v2.0`,
          algorithms: ['RS256']
        }) as TokenPayload;

        resolve(payload);
      } catch (error) {
        reject(error);
      }
    });
  }

  createMiddleware(requiredScopes: string[] = []) {
    return async (req: Request, res: Response, next: NextFunction) => {
      try {
        const authHeader = req.headers.authorization;
        
        if (!authHeader || !authHeader.startsWith('Bearer ')) {
          res.setHeader('WWW-Authenticate', 'Bearer realm="outlook-mcp"');
          return res.status(401).json({
            error: 'unauthorized',
            error_description: 'Bearer token required'
          });
        }

        const token = authHeader.substring(7);
        const payload = await this.validateToken(token);

        // Check required scopes
        if (requiredScopes.length > 0) {
          const tokenScopes = payload.scp ? payload.scp.split(' ') : [];
          const hasRequiredScopes = requiredScopes.every(scope => 
            tokenScopes.includes(scope)
          );

          if (!hasRequiredScopes) {
            return res.status(403).json({
              error: 'insufficient_scope',
              error_description: `Required scopes: ${requiredScopes.join(', ')}`
            });
          }
        }

        // Add token info to request
        (req as any).tokenPayload = payload;
        next();
      } catch (error) {
        res.setHeader('WWW-Authenticate', 'Bearer realm="outlook-mcp", error="invalid_token"');
        return res.status(401).json({
          error: 'invalid_token',
          error_description: error instanceof Error ? error.message : 'Token validation failed'
        });
      }
    };
  }
}

export const tokenValidator = new MicrosoftTokenValidator();