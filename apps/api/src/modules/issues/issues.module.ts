import { Module } from '@nestjs/common';
import { AuthModule } from '../auth/auth.module';
import { AuditLogsModule } from '../audit-logs/audit-logs.module';
import { FeatureFlagsModule } from '../feature-flags/feature-flags.module';
import { IssuesController } from './issues.controller';
import { IssuesService } from './issues.service';

@Module({
  imports: [AuthModule, FeatureFlagsModule, AuditLogsModule],
  controllers: [IssuesController],
  providers: [IssuesService],
})
export class IssuesModule {}
