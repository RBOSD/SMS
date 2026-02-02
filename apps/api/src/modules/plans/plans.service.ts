import {
  BadRequestException,
  ConflictException,
  Injectable,
  NotFoundException,
} from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';

function parseDateOrNull(v?: string): Date | null {
  if (!v) return null;
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  return Number.isFinite(d.getTime()) ? d : null;
}

@Injectable()
export class PlansService {
  constructor(private readonly prisma: PrismaService) {}

  async list(params: {
    page: number;
    pageSize: number;
    q?: string;
    year?: string;
  }) {
    const page = Math.max(params.page || 1, 1);
    const pageSize = Math.min(Math.max(params.pageSize || 20, 1), 200);
    const skip = (page - 1) * pageSize;
    const take = pageSize;
    const q = String(params.q || '').trim();
    const year = String(params.year || '').trim();

    const where: any = {};
    if (year) where.year = year;
    if (q) where.name = { contains: q, mode: 'insensitive' as const };

    const [total, rows] = await Promise.all([
      this.prisma.plan.count({ where }),
      this.prisma.plan.findMany({
        where,
        orderBy: [{ year: 'desc' }, { name: 'asc' }],
        skip,
        take,
      }),
    ]);
    return { page, pageSize, total, rows };
  }

  async create(params: { name: string; year: string; status?: string }) {
    const name = String(params.name || '').trim();
    const year = String(params.year || '').trim();
    if (!name || !year) throw new BadRequestException('name/year required');

    try {
      return await this.prisma.plan.create({
        data: { name, year, status: params.status?.trim() || null },
      });
    } catch (e: any) {
      if (String(e?.code) === 'P2002')
        throw new ConflictException('plan already exists');
      throw e;
    }
  }

  async listSchedules(planId: number) {
    if (!Number.isFinite(planId))
      throw new BadRequestException('invalid planId');
    const plan = await this.prisma.plan.findUnique({ where: { id: planId } });
    if (!plan) throw new NotFoundException('plan not found');
    const rows = await this.prisma.planSchedule.findMany({
      where: { planId },
      orderBy: [{ startDate: 'asc' }, { id: 'asc' }],
    });
    return { plan, rows };
  }

  async createSchedule(
    planId: number,
    params: {
      railway: string;
      inspectionType: string;
      planNumber?: string;
      startDate?: string;
      endDate?: string;
      business?: string;
      location?: string;
      inspector?: string;
    },
  ) {
    if (!Number.isFinite(planId))
      throw new BadRequestException('invalid planId');
    const plan = await this.prisma.plan.findUnique({ where: { id: planId } });
    if (!plan) throw new NotFoundException('plan not found');

    const railway = String(params.railway || '').trim();
    const inspectionType = String(params.inspectionType || '').trim();
    if (!railway || !inspectionType)
      throw new BadRequestException('railway/inspectionType required');

    return await this.prisma.planSchedule.create({
      data: {
        planId,
        railway,
        inspectionType,
        planNumber: params.planNumber?.trim() || null,
        startDate: parseDateOrNull(params.startDate),
        endDate: parseDateOrNull(params.endDate),
        business: params.business?.trim() || null,
        location: params.location?.trim() || null,
        inspector: params.inspector?.trim() || null,
      },
    });
  }
}
