import { ForbiddenException, Injectable } from '@nestjs/common';
import { JwtService } from '@nestjs/jwt';
import type { Role } from '@prisma/client';
import * as bcrypt from 'bcrypt';
import { PrismaService } from '../prisma/prisma.service';

@Injectable()
export class AuthService {
  constructor(
    private readonly prisma: PrismaService,
    private readonly jwt: JwtService,
  ) {}

  async validateUser(username: string, password: string) {
    const u = await this.prisma.user.findUnique({
      where: { username },
    });
    if (!u) return null;
    const ok = await bcrypt.compare(password, u.passwordHash);
    if (!ok) return null;
    return u;
  }

  async isAdminUser(userId: number): Promise<boolean> {
    const r = await this.prisma.userGroup.findFirst({
      where: {
        userId,
        group: { isAdminGroup: true },
      },
      select: { userId: true },
    });
    return !!r;
  }

  async signToken(payload: { sub: number; username: string; role: Role }) {
    return await this.jwt.signAsync(payload);
  }

  async requireCanLogin(userId: number) {
    // placeholder for future account lockout/policy checks
    if (!userId) throw new ForbiddenException();
  }
}
