import type { Role } from '@prisma/client';

export interface JwtPayload {
  sub: number;
  username: string;
  role: Role;
}

export interface RequestUser extends JwtPayload {
  // alias for convenience
  userId: number;
}

export function toRequestUser(payload: JwtPayload): RequestUser {
  return { ...payload, userId: payload.sub };
}

export function getRequestIp(req: any): string | null {
  const xf = req?.headers?.['x-forwarded-for'];
  if (typeof xf === 'string' && xf.trim()) return xf.split(',')[0].trim();
  const ra = req?.socket?.remoteAddress;
  return typeof ra === 'string' ? ra : null;
}
