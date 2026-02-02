import { BadRequestException, Injectable } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';

@Injectable()
export class GroupsService {
  constructor(private readonly prisma: PrismaService) {}

  async list() {
    return await this.prisma.group.findMany({
      orderBy: [{ isAdminGroup: 'desc' }, { name: 'asc' }],
    });
  }

  async create(params: { name: string; isAdminGroup?: boolean }) {
    const name = String(params.name || '').trim();
    if (!name) throw new BadRequestException('name is required');

    const isAdminGroup = params.isAdminGroup === true;
    if (isAdminGroup) {
      // Ensure only one admin group
      await this.prisma.group.updateMany({
        where: { isAdminGroup: true },
        data: { isAdminGroup: false },
      });
    }

    return await this.prisma.group.create({
      data: { name, isAdminGroup },
    });
  }

  async update(id: number, params: { name?: string; isAdminGroup?: boolean }) {
    if (!Number.isFinite(id)) throw new BadRequestException('invalid id');

    const data: any = {};
    if (params.name != null) {
      const name = String(params.name || '').trim();
      if (!name) throw new BadRequestException('name is required');
      data.name = name;
    }
    if (params.isAdminGroup != null) {
      const isAdminGroup = params.isAdminGroup === true;
      if (isAdminGroup) {
        await this.prisma.group.updateMany({
          where: { isAdminGroup: true },
          data: { isAdminGroup: false },
        });
      }
      data.isAdminGroup = isAdminGroup;
    }

    return await this.prisma.group.update({
      where: { id },
      data,
    });
  }
}
