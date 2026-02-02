import {
  CanActivate,
  ExecutionContext,
  Injectable,
  UnauthorizedException,
} from '@nestjs/common';
import { Role } from '@prisma/client';
import { AuthService } from '../modules/auth/auth.service';

@Injectable()
export class AdminOrManagerGuard implements CanActivate {
  constructor(private readonly auth: AuthService) {}

  async canActivate(ctx: ExecutionContext): Promise<boolean> {
    const req = ctx.switchToHttp().getRequest();
    const user = req.user as { userId?: number; role?: Role } | undefined;
    if (!user?.userId) throw new UnauthorizedException();
    if (user.role === Role.MANAGER) return true;
    return await this.auth.isAdminUser(user.userId);
  }
}
