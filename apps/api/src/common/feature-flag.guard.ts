import {
  CanActivate,
  ExecutionContext,
  Injectable,
  NotFoundException,
  UnauthorizedException,
} from '@nestjs/common';
import { Reflector } from '@nestjs/core';
import { FEATURE_FLAG_KEY } from './feature-flag.decorator';
import { FeatureFlagsService } from '../modules/feature-flags/feature-flags.service';

@Injectable()
export class FeatureFlagGuard implements CanActivate {
  constructor(
    private readonly reflector: Reflector,
    private readonly features: FeatureFlagsService,
  ) {}

  async canActivate(ctx: ExecutionContext): Promise<boolean> {
    const key = this.reflector.getAllAndOverride<string>(FEATURE_FLAG_KEY, [
      ctx.getHandler(),
      ctx.getClass(),
    ]);
    if (!key) return true;

    const req = ctx.switchToHttp().getRequest();
    const user = req.user as { userId?: number } | undefined;
    if (!user?.userId) throw new UnauthorizedException();

    const ok = await this.features.isEnabledForUser(user.userId, key);
    if (!ok) throw new NotFoundException();
    return true;
  }
}
