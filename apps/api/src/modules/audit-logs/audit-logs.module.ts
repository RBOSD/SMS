import { Module, forwardRef } from '@nestjs/common';
import { AuthModule } from '../auth/auth.module';
import { AuditLogsAdminController } from './audit-logs.controller';
import { AuditLogsService } from './audit-logs.service';

@Module({
  imports: [forwardRef(() => AuthModule)],
  controllers: [AuditLogsAdminController],
  providers: [AuditLogsService],
  exports: [AuditLogsService],
})
export class AuditLogsModule { }
