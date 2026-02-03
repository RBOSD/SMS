import { Module, forwardRef } from '@nestjs/common';
import { AuditLogsModule } from '../audit-logs/audit-logs.module';
import { AuthModule } from '../auth/auth.module';
import { FeatureFlagGuard } from '../../common/feature-flag.guard';
import { FeatureFlagsAdminController } from './feature-flags.controller';
import { FeatureFlagsService } from './feature-flags.service';

@Module({
  imports: [forwardRef(() => AuthModule), forwardRef(() => AuditLogsModule)],
  controllers: [FeatureFlagsAdminController],
  providers: [FeatureFlagsService, FeatureFlagGuard],
  exports: [FeatureFlagsService, FeatureFlagGuard],
})
export class FeatureFlagsModule { }
