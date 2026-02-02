import {
  BadRequestException,
  ConflictException,
  Injectable,
  NotFoundException,
} from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';

@Injectable()
export class IssuesService {
  constructor(private readonly prisma: PrismaService) {}

  async list(params: {
    page: number;
    pageSize: number;
    q?: string;
    year?: string;
    planId?: number;
  }) {
    const page = Math.max(params.page || 1, 1);
    const pageSize = Math.min(Math.max(params.pageSize || 20, 1), 200);
    const skip = (page - 1) * pageSize;
    const take = pageSize;
    const q = String(params.q || '').trim();
    const year = String(params.year || '').trim();
    const planId = params.planId;

    const where: any = {};
    if (year) where.year = year;
    if (Number.isFinite(planId as any)) where.planId = planId;
    if (q) {
      where.OR = [
        { number: { contains: q, mode: 'insensitive' as const } },
        { unit: { contains: q, mode: 'insensitive' as const } },
        { content: { contains: q, mode: 'insensitive' as const } },
      ];
    }

    const [total, rows] = await Promise.all([
      this.prisma.issue.count({ where }),
      this.prisma.issue.findMany({
        where,
        orderBy: { updatedAt: 'desc' },
        skip,
        take,
        include: {
          plan: true,
          rounds: { orderBy: { round: 'desc' }, take: 1 },
        },
      }),
    ]);

    return { page, pageSize, total, rows };
  }

  async create(params: {
    number: string;
    year?: string;
    unit?: string;
    content?: string;
    status?: string;
    planId?: number;
  }) {
    const number = String(params.number || '').trim();
    if (!number) throw new BadRequestException('number is required');

    try {
      return await this.prisma.issue.create({
        data: {
          number,
          year: params.year?.trim() || null,
          unit: params.unit?.trim() || null,
          content: params.content || null,
          status: params.status?.trim() || null,
          planId: Number.isFinite(params.planId as any) ? params.planId : null,
        },
      });
    } catch (e: any) {
      if (String(e?.code) === 'P2002')
        throw new ConflictException('number already exists');
      throw e;
    }
  }

  async upsertRound(
    issueId: number,
    params: {
      round: number;
      handling?: string;
      review?: string;
      replyDate?: string;
      responseDate?: string;
    },
  ) {
    if (!Number.isFinite(issueId))
      throw new BadRequestException('invalid issueId');
    const round = Math.max(parseInt(String(params.round || ''), 10) || 0, 0);
    if (round < 1) throw new BadRequestException('round must be >= 1');

    const issue = await this.prisma.issue.findUnique({
      where: { id: issueId },
    });
    if (!issue) throw new NotFoundException('issue not found');

    return await this.prisma.issueRound.upsert({
      where: { issueId_round: { issueId, round } },
      update: {
        handling: params.handling ?? undefined,
        review: params.review ?? undefined,
        replyDate: params.replyDate ?? undefined,
        responseDate: params.responseDate ?? undefined,
      },
      create: {
        issueId,
        round,
        handling: params.handling ?? null,
        review: params.review ?? null,
        replyDate: params.replyDate ?? null,
        responseDate: params.responseDate ?? null,
      },
    });
  }
}
