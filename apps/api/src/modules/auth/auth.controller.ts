import {
  Body,
  Controller,
  Get,
  Post,
  Req,
  Res,
  UnauthorizedException,
  UseGuards,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import type { Response } from 'express';
import { getRequestIp } from '../../common/request-user';
import type { RequestUser } from '../../common/request-user';
import { JwtAuthGuard } from './jwt-auth.guard';
import { AuthService } from './auth.service';
import { LoginDto } from './dto/login.dto';
import { FeatureFlagsService } from '../feature-flags/feature-flags.service';

@Controller('auth')
export class AuthController {
  constructor(
    private readonly auth: AuthService,
    private readonly config: ConfigService,
    private readonly features: FeatureFlagsService,
  ) {}

  @Post('login')
  async login(@Body() body: LoginDto, @Req() req: any, @Res() res: Response) {
    const username = String(body.username || '').trim();
    const password = String(body.password || '');
    const u = await this.auth.validateUser(username, password);
    if (!u) throw new UnauthorizedException('Invalid username or password');

    await this.auth.requireCanLogin(u.id);

    const token = await this.auth.signToken({
      sub: u.id,
      username: u.username,
      role: u.role,
    });

    const cookieName =
      this.config.get<string>('JWT_COOKIE_NAME') || 'sms_token';

    res.cookie(cookieName, token, {
      httpOnly: true,
      sameSite: 'lax',
      secure:
        (this.config.get<string>('NODE_ENV') || process.env.NODE_ENV) ===
        'production',
      path: '/',
      maxAge: 7 * 24 * 60 * 60 * 1000,
    });

    const isAdmin = await this.auth.isAdminUser(u.id);
    const features = await this.features.getEffectiveForUser(u.id);

    return res.json({
      isLogin: true,
      id: u.id,
      username: u.username,
      name: u.name,
      role: u.role,
      isAdmin,
      features,
      ip: getRequestIp(req),
    });
  }

  @Post('logout')
  async logout(@Res() res: Response) {
    const cookieName =
      this.config.get<string>('JWT_COOKIE_NAME') || 'sms_token';
    res.clearCookie(cookieName, { path: '/' });
    return res.json({ success: true });
  }

  @Get('me')
  @UseGuards(JwtAuthGuard)
  async me(@Req() req: any) {
    const u = req.user as RequestUser | undefined;
    if (!u) throw new UnauthorizedException();
    const isAdmin = await this.auth.isAdminUser(u.userId);
    const features = await this.features.getEffectiveForUser(u.userId);
    return {
      isLogin: true,
      id: u.userId,
      username: u.username,
      role: u.role,
      isAdmin,
      features,
    };
  }
}
