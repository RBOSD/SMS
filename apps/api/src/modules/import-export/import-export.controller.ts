import {
  Controller,
  Get,
  Post,
  Res,
  UseGuards,
  UploadedFile,
  UseInterceptors,
  BadRequestException,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import type { Response } from 'express';
import * as Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AdminGuard } from '../../common/admin.guard';
import { PrismaService } from '../prisma/prisma.service';

@Controller('admin')
@UseGuards(JwtAuthGuard, AdminGuard)
export class ImportExportAdminController {
  constructor(private readonly prisma: PrismaService) {}

  @Get('export/issues.csv')
  async exportIssuesCsv(@Res() res: Response) {
    const rows = await this.prisma.issue.findMany({
      orderBy: { updatedAt: 'desc' },
      include: { plan: true },
    });

    const csv = Papa.unparse(
      rows.map((r) => ({
        number: r.number,
        year: r.year ?? '',
        unit: r.unit ?? '',
        status: r.status ?? '',
        content: r.content ?? '',
        planName: r.plan?.name ?? '',
        planYear: r.plan?.year ?? '',
      })),
    );

    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="issues.csv"');
    return res.send('\uFEFF' + csv);
  }

  @Get('export/plans.csv')
  async exportPlansCsv(@Res() res: Response) {
    const rows = await this.prisma.plan.findMany({
      orderBy: [{ year: 'desc' }, { name: 'asc' }],
    });

    const csv = Papa.unparse(
      rows.map((r) => ({
        name: r.name,
        year: r.year,
        status: r.status ?? '',
      })),
    );

    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="plans.csv"');
    return res.send('\uFEFF' + csv);
  }

  @Post('import/issues')
  @UseInterceptors(FileInterceptor('file'))
  async importIssues(@UploadedFile() file?: Express.Multer.File) {
    if (!file?.buffer) throw new BadRequestException('file is required');
    const name = String(file.originalname || '').toLowerCase();

    let records: Record<string, any>[] = [];
    if (name.endsWith('.csv')) {
      const text = file.buffer.toString('utf8');
      const parsed = Papa.parse<Record<string, any>>(text, {
        header: true,
        skipEmptyLines: true,
      });
      if (parsed.errors?.length) {
        throw new BadRequestException(
          parsed.errors[0]?.message || 'CSV parse error',
        );
      }
      records = parsed.data;
    } else if (name.endsWith('.xlsx')) {
      const wb = XLSX.read(file.buffer, { type: 'buffer' });
      const sheetName = wb.SheetNames?.[0];
      if (!sheetName) throw new BadRequestException('XLSX has no sheets');
      const sheet = wb.Sheets[sheetName];
      records = XLSX.utils.sheet_to_json<Record<string, any>>(sheet, {
        defval: '',
      });
    } else {
      throw new BadRequestException('only .csv or .xlsx supported');
    }

    let created = 0;
    let updated = 0;
    let skipped = 0;

    for (const r of records) {
      const number = String(r.number || r['編號'] || '').trim();
      if (!number) {
        skipped++;
        continue;
      }
      const year = String(r.year || r['年度'] || '').trim() || null;
      const unit = String(r.unit || r['機構'] || '').trim() || null;
      const status = String(r.status || r['狀態'] || '').trim() || null;
      const content = String(r.content || r['事項內容'] || '').trim() || null;

      const existed = await this.prisma.issue.findUnique({ where: { number } });
      await this.prisma.issue.upsert({
        where: { number },
        update: { year, unit, status, content },
        create: { number, year, unit, status, content },
      });
      if (existed) updated++;
      else created++;
    }

    return { success: true, created, updated, skipped };
  }

  @Post('import/plans')
  @UseInterceptors(FileInterceptor('file'))
  async importPlans(@UploadedFile() file?: Express.Multer.File) {
    if (!file?.buffer) throw new BadRequestException('file is required');
    const name = String(file.originalname || '').toLowerCase();

    let records: Record<string, any>[] = [];
    if (name.endsWith('.csv')) {
      const text = file.buffer.toString('utf8');
      const parsed = Papa.parse<Record<string, any>>(text, {
        header: true,
        skipEmptyLines: true,
      });
      if (parsed.errors?.length) {
        throw new BadRequestException(
          parsed.errors[0]?.message || 'CSV parse error',
        );
      }
      records = parsed.data;
    } else if (name.endsWith('.xlsx')) {
      const wb = XLSX.read(file.buffer, { type: 'buffer' });
      const sheetName = wb.SheetNames?.[0];
      if (!sheetName) throw new BadRequestException('XLSX has no sheets');
      const sheet = wb.Sheets[sheetName];
      records = XLSX.utils.sheet_to_json<Record<string, any>>(sheet, {
        defval: '',
      });
    } else {
      throw new BadRequestException('only .csv or .xlsx supported');
    }

    let created = 0;
    let updated = 0;
    let skipped = 0;

    for (const r of records) {
      const name = String(r.name || r['計畫名稱'] || '').trim();
      const year = String(r.year || r['年度'] || '').trim();
      if (!name || !year) {
        skipped++;
        continue;
      }
      const status = String(r.status || r['狀態'] || '').trim() || null;

      const existed = await this.prisma.plan.findUnique({
        where: { name_year: { name, year } },
      });
      await this.prisma.plan.upsert({
        where: { name_year: { name, year } },
        update: { status },
        create: { name, year, status },
      });
      if (existed) updated++;
      else created++;
    }

    return { success: true, created, updated, skipped };
  }
}
