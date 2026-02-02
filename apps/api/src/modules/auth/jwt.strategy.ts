import { Injectable, UnauthorizedException } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { PassportStrategy } from '@nestjs/passport';
import { ExtractJwt, Strategy } from 'passport-jwt';
import type { Request } from 'express';
import type { JwtPayload } from '../../common/request-user';
import { toRequestUser } from '../../common/request-user';

function cookieExtractor(req: Request, cookieName: string): string | null {
  const cookies = (req as any)?.cookies;
  if (!cookies) return null;
  const v = cookies[cookieName];
  return typeof v === 'string' && v.trim() ? v : null;
}

@Injectable()
export class JwtStrategy extends PassportStrategy(Strategy) {
  constructor(private readonly config: ConfigService) {
    const cookieName = config.get<string>('JWT_COOKIE_NAME') || 'sms_token';

    super({
      jwtFromRequest: ExtractJwt.fromExtractors([
        (req: Request) => cookieExtractor(req, cookieName),
        ExtractJwt.fromAuthHeaderAsBearerToken(),
      ]),
      ignoreExpiration: false,
      secretOrKey: config.get<string>('JWT_SECRET') || 'dev-only-change-me',
    });
  }

  validate(payload: JwtPayload) {
    if (!payload?.sub || !payload.username) throw new UnauthorizedException();
    return toRequestUser(payload);
  }
}
