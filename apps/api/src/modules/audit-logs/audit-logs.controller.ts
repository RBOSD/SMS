import { Controller, Get, Query, UseGuards } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AdminGuard } from '../../common/admin.guard';

@Controller('admin/audit-logs')
@UseGuards(JwtAuthGuard, AdminGuard)
export class AuditLogsAdminController {
  constructor(private readonly prisma: PrismaService) {}

  @Get()
  async list(
    @Query('page') pageRaw?: string,
    @Query('pageSize') pageSizeRaw?: string,
    @Query('q') q?: string,
  ) {
    const page = Math.max(parseInt(pageRaw || '1', 10) || 1, 1);
    const pageSize = Math.min(
      Math.max(parseInt(pageSizeRaw || '20', 10) || 20, 1),
      200,
    );
    const skip = (page - 1) * pageSize;
    const take = pageSize;
    const keyword = String(q || '').trim();

    const where = keyword
      ? {
          OR: [
            { action: { contains: keyword, mode: 'insensitive' as const } },
            { details: { contains: keyword, mode: 'insensitive' as const } },
            { ip: { contains: keyword, mode: 'insensitive' as const } },
          ],
        }
      : {};

    const [total, rows] = await Promise.all([
      this.prisma.auditLog.count({ where }),
      this.prisma.auditLog.findMany({
        where,
        orderBy: { createdAt: 'desc' },
        skip,
        take,
        include: {
          actor: { select: { id: true, username: true, name: true } },
        },
      }),
    ]);

    return {
      page,
      pageSize,
      total,
      data: rows.map((r) => ({
        id: r.id,
        createdAt: r.createdAt,
        action: r.action,
        details: r.details,
        ip: r.ip,
        actor: r.actor
          ? { id: r.actor.id, username: r.actor.username, name: r.actor.name }
          : null,
      })),
    };
  }
}
