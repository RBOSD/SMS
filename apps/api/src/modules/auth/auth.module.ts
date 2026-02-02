import { Module, forwardRef } from '@nestjs/common';
import { ConfigModule, ConfigService } from '@nestjs/config';
import { JwtModule } from '@nestjs/jwt';
import { PassportModule } from '@nestjs/passport';
import { AuthController } from './auth.controller';
import { AuthService } from './auth.service';
import { JwtStrategy } from './jwt.strategy';
import { AdminGuard } from '../../common/admin.guard';
import { AdminOrManagerGuard } from '../../common/admin-or-manager.guard';
import { FeatureFlagsModule } from '../feature-flags/feature-flags.module';

@Module({
  imports: [
    ConfigModule,
    PassportModule,
    JwtModule.registerAsync({
      imports: [ConfigModule],
      inject: [ConfigService],
      useFactory: (config: ConfigService) => ({
        secret: config.get<string>('JWT_SECRET') || 'dev-only-change-me',
        signOptions: { expiresIn: '7d' },
      }),
    }),
    forwardRef(() => FeatureFlagsModule),
  ],
  controllers: [AuthController],
  providers: [AuthService, JwtStrategy, AdminGuard, AdminOrManagerGuard],
  exports: [AuthService, AdminGuard, AdminOrManagerGuard],
})
export class AuthModule {}
