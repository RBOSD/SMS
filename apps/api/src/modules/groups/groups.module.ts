import { Module } from '@nestjs/common';
import { AuthModule } from '../auth/auth.module';
import { GroupsAdminController } from './groups.controller';
import { GroupsService } from './groups.service';

@Module({
  imports: [AuthModule],
  controllers: [GroupsAdminController],
  providers: [GroupsService],
  exports: [GroupsService],
})
export class GroupsModule {}
