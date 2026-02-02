import {
  Body,
  Controller,
  Get,
  Param,
  Post,
  Put,
  Query,
  Req,
  UseGuards,
} from '@nestjs/common';
import { RequireFeature } from '../../common/feature-flag.decorator';
import { FeatureFlagGuard } from '../../common/feature-flag.guard';
import { AdminOrManagerGuard } from '../../common/admin-or-manager.guard';
import { getRequestIp } from '../../common/request-user';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AuditLogsService } from '../audit-logs/audit-logs.service';
import { CreateIssueDto } from './dto/create-issue.dto';
import { UpsertIssueRoundDto } from './dto/upsert-issue-round.dto';
import { IssuesService } from './issues.service';

@Controller('issues')
@RequireFeature('module_issues')
export class IssuesController {
  constructor(
    private readonly issues: IssuesService,
    private readonly audit: AuditLogsService,
  ) {}

  @Get()
  @UseGuards(JwtAuthGuard, FeatureFlagGuard)
  async list(
    @Query('page') pageRaw?: string,
    @Query('pageSize') pageSizeRaw?: string,
    @Query('q') q?: string,
    @Query('year') year?: string,
    @Query('planId') planIdRaw?: string,
  ) {
    const page = parseInt(pageRaw || '1', 10) || 1;
    const pageSize = parseInt(pageSizeRaw || '20', 10) || 20;
    const planId = planIdRaw != null ? parseInt(planIdRaw, 10) : undefined;

    const r = await this.issues.list({ page, pageSize, q, year, planId });
    return {
      page: r.page,
      pageSize: r.pageSize,
      total: r.total,
      data: r.rows.map((x) => ({
        id: x.id,
        number: x.number,
        year: x.year,
        unit: x.unit,
        status: x.status,
        content: x.content,
        plan: x.plan
          ? { id: x.plan.id, name: x.plan.name, year: x.plan.year }
          : null,
        latestRound: x.rounds?.[0]
          ? {
              round: x.rounds[0].round,
              handling: x.rounds[0].handling,
              review: x.rounds[0].review,
              replyDate: x.rounds[0].replyDate,
              responseDate: x.rounds[0].responseDate,
            }
          : null,
        createdAt: x.createdAt,
        updatedAt: x.updatedAt,
      })),
    };
  }

  @Post()
  @UseGuards(JwtAuthGuard, FeatureFlagGuard, AdminOrManagerGuard)
  async create(@Body() body: CreateIssueDto, @Req() req: any) {
    const r = await this.issues.create({
      number: body.number,
      year: body.year,
      unit: body.unit,
      content: body.content,
      status: body.status,
      planId: body.planId,
    });

    await this.audit.log({
      actorUserId: req.user?.userId,
      action: 'CREATE_ISSUE',
      details: `create issue number=${r.number} id=${r.id}`,
      ip: getRequestIp(req),
    });

    return { data: r };
  }

  @Put(':id/rounds')
  @UseGuards(JwtAuthGuard, FeatureFlagGuard, AdminOrManagerGuard)
  async upsertRound(
    @Param('id') idRaw: string,
    @Body() body: UpsertIssueRoundDto,
    @Req() req: any,
  ) {
    const id = parseInt(idRaw, 10);
    const r = await this.issues.upsertRound(id, {
      round: body.round,
      handling: body.handling,
      review: body.review,
      replyDate: body.replyDate,
      responseDate: body.responseDate,
    });

    await this.audit.log({
      actorUserId: req.user?.userId,
      action: 'UPSERT_ISSUE_ROUND',
      details: `issueId=${id} round=${body.round}`,
      ip: getRequestIp(req),
    });

    return { data: r };
  }
}
