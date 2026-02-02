import {
  Body,
  Controller,
  Get,
  Param,
  Post,
  Query,
  Req,
  UseGuards,
} from '@nestjs/common';
import { AdminOrManagerGuard } from '../../common/admin-or-manager.guard';
import { RequireFeature } from '../../common/feature-flag.decorator';
import { FeatureFlagGuard } from '../../common/feature-flag.guard';
import { getRequestIp } from '../../common/request-user';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AuditLogsService } from '../audit-logs/audit-logs.service';
import { CreatePlanDto } from './dto/create-plan.dto';
import { CreatePlanScheduleDto } from './dto/create-plan-schedule.dto';
import { PlansService } from './plans.service';

@Controller('plans')
@RequireFeature('module_plans')
export class PlansController {
  constructor(
    private readonly plans: PlansService,
    private readonly audit: AuditLogsService,
  ) {}

  @Get()
  @UseGuards(JwtAuthGuard, FeatureFlagGuard)
  async list(
    @Query('page') pageRaw?: string,
    @Query('pageSize') pageSizeRaw?: string,
    @Query('q') q?: string,
    @Query('year') year?: string,
  ) {
    const page = parseInt(pageRaw || '1', 10) || 1;
    const pageSize = parseInt(pageSizeRaw || '20', 10) || 20;
    const r = await this.plans.list({ page, pageSize, q, year });
    return {
      page: r.page,
      pageSize: r.pageSize,
      total: r.total,
      data: r.rows,
    };
  }

  @Post()
  @UseGuards(JwtAuthGuard, FeatureFlagGuard, AdminOrManagerGuard)
  async create(@Body() body: CreatePlanDto, @Req() req: any) {
    const p = await this.plans.create({
      name: body.name,
      year: body.year,
      status: body.status,
    });

    await this.audit.log({
      actorUserId: req.user?.userId,
      action: 'CREATE_PLAN',
      details: `create plan id=${p.id} ${p.name}(${p.year})`,
      ip: getRequestIp(req),
    });

    return { data: p };
  }

  @Get(':id/schedules')
  @UseGuards(JwtAuthGuard, FeatureFlagGuard)
  async listSchedules(@Param('id') idRaw: string) {
    const id = parseInt(idRaw, 10);
    const r = await this.plans.listSchedules(id);
    return {
      plan: r.plan,
      data: r.rows,
    };
  }

  @Post(':id/schedules')
  @UseGuards(JwtAuthGuard, FeatureFlagGuard, AdminOrManagerGuard)
  async createSchedule(
    @Param('id') idRaw: string,
    @Body() body: CreatePlanScheduleDto,
    @Req() req: any,
  ) {
    const id = parseInt(idRaw, 10);
    const s = await this.plans.createSchedule(id, body);
    await this.audit.log({
      actorUserId: req.user?.userId,
      action: 'CREATE_PLAN_SCHEDULE',
      details: `create plan schedule planId=${id} scheduleId=${s.id}`,
      ip: getRequestIp(req),
    });
    return { data: s };
  }
}
