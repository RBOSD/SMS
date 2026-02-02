import { Injectable } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';

@Injectable()
export class AuditLogsService {
  constructor(private readonly prisma: PrismaService) {}

  async log(params: {
    actorUserId?: number | null;
    action: string;
    details?: string | null;
    ip?: string | null;
  }) {
    await this.prisma.auditLog.create({
      data: {
        actorUserId: params.actorUserId ?? null,
        action: params.action,
        details: params.details ?? null,
        ip: params.ip ?? null,
      },
    });
  }
}
