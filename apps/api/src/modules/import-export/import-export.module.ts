import { Module } from '@nestjs/common';
import { AuthModule } from '../auth/auth.module';
import { ImportExportAdminController } from './import-export.controller';

@Module({
  imports: [AuthModule],
  controllers: [ImportExportAdminController],
})
export class ImportExportModule {}
