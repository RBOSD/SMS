import { Module } from '@nestjs/common';
import { AuthModule } from '../auth/auth.module';
import { AuditLogsModule } from '../audit-logs/audit-logs.module';
import { FeatureFlagsModule } from '../feature-flags/feature-flags.module';
import { PlansController } from './plans.controller';
import { PlansService } from './plans.service';

@Module({
  imports: [AuthModule, FeatureFlagsModule, AuditLogsModule],
  controllers: [PlansController],
  providers: [PlansService],
})
export class PlansModule {}
