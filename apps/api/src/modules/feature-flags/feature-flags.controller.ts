import { Body, Controller, Get, Put, Req, UseGuards } from '@nestjs/common';
import { AdminGuard } from '../../common/admin.guard';
import { getRequestIp } from '../../common/request-user';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AuditLogsService } from '../audit-logs/audit-logs.service';
import { UpdateFeatureFlagsDto } from './dto/update-feature-flags.dto';
import { FeatureFlagsService } from './feature-flags.service';

@Controller('admin/feature-flags')
@UseGuards(JwtAuthGuard, AdminGuard)
export class FeatureFlagsAdminController {
  constructor(
    private readonly features: FeatureFlagsService,
    private readonly audit: AuditLogsService,
  ) {}

  @Get()
  async getAll() {
    await this.features.ensureSeed();
    return { data: await this.features.getAll() };
  }

  @Put()
  async update(@Body() body: UpdateFeatureFlagsDto, @Req() req: any) {
    await this.features.ensureSeed();
    const updated = await this.features.update(body.flags || {});
    await this.audit.log({
      actorUserId: req.user?.userId,
      action: 'UPDATE_FEATURE_FLAGS',
      details: JSON.stringify(updated),
      ip: getRequestIp(req),
    });
    return { data: updated };
  }
}
