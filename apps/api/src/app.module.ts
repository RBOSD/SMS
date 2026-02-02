import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { AppInitService } from './app-init.service';
import { FeatureFlagGuard } from './common/feature-flag.guard';
import { RolesGuard } from './common/roles.guard';
import { AuthModule } from './modules/auth/auth.module';
import { AuditLogsModule } from './modules/audit-logs/audit-logs.module';
import { AiReviewModule } from './modules/ai-review/ai-review.module';
import { FeatureFlagsModule } from './modules/feature-flags/feature-flags.module';
import { GroupsModule } from './modules/groups/groups.module';
import { IssuesModule } from './modules/issues/issues.module';
import { ImportExportModule } from './modules/import-export/import-export.module';
import { PlansModule } from './modules/plans/plans.module';
import { PrismaModule } from './modules/prisma/prisma.module';
import { UsersModule } from './modules/users/users.module';

@Module({
  imports: [
    ConfigModule.forRoot({
      isGlobal: true,
      envFilePath: ['.env'],
    }),
    PrismaModule,
    AuthModule,
    AuditLogsModule,
    AiReviewModule,
    FeatureFlagsModule,
    GroupsModule,
    UsersModule,
    IssuesModule,
    PlansModule,
    ImportExportModule,
  ],
  controllers: [AppController],
  providers: [AppService, AppInitService, RolesGuard, FeatureFlagGuard],
})
export class AppModule {}
