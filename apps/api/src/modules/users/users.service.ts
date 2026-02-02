import {
  BadRequestException,
  ConflictException,
  Injectable,
  NotFoundException,
} from '@nestjs/common';
import { Role } from '@prisma/client';
import * as bcrypt from 'bcrypt';
import { PrismaService } from '../prisma/prisma.service';

@Injectable()
export class UsersService {
  constructor(private readonly prisma: PrismaService) {}

  async list() {
    return await this.prisma.user.findMany({
      orderBy: [{ createdAt: 'desc' }],
      include: {
        groups: {
          include: { group: true },
        },
      },
    });
  }

  async create(params: {
    username: string;
    password: string;
    name?: string | null;
    role?: Role;
    groupIds?: number[];
  }) {
    const username = String(params.username || '').trim();
    const password = String(params.password || '');
    if (!username) throw new BadRequestException('username is required');
    if (password.length < 8)
      throw new BadRequestException('password too short');

    const role = params.role ?? Role.VIEWER;
    const passwordHash = await bcrypt.hash(password, 10);

    try {
      const u = await this.prisma.user.create({
        data: {
          username,
          passwordHash,
          name: params.name?.trim() || null,
          role,
          mustChangePassword: true,
        },
      });

      const groupIds = (params.groupIds || []).filter((n) =>
        Number.isFinite(n),
      );
      if (groupIds.length > 0) {
        await this.prisma.userGroup.createMany({
          data: groupIds.map((gid) => ({ userId: u.id, groupId: gid })),
          skipDuplicates: true,
        });
      }

      return await this.prisma.user.findUniqueOrThrow({
        where: { id: u.id },
        include: { groups: { include: { group: true } } },
      });
    } catch (e: any) {
      if (String(e?.code) === 'P2002')
        throw new ConflictException('username already exists');
      throw e;
    }
  }

  async update(
    id: number,
    params: {
      name?: string;
      password?: string;
      role?: Role;
      groupIds?: number[];
    },
  ) {
    if (!Number.isFinite(id)) throw new BadRequestException('invalid id');
    const existing = await this.prisma.user.findUnique({ where: { id } });
    if (!existing) throw new NotFoundException('user not found');

    const data: any = {};
    if (params.name != null)
      data.name = String(params.name || '').trim() || null;
    if (params.role != null) data.role = params.role;
    if (params.password != null) {
      if (params.password.length < 8)
        throw new BadRequestException('password too short');
      data.passwordHash = await bcrypt.hash(params.password, 10);
      data.mustChangePassword = true;
    }

    await this.prisma.user.update({ where: { id }, data });

    if (params.groupIds != null) {
      const groupIds = (params.groupIds || []).filter((n) =>
        Number.isFinite(n),
      );
      await this.prisma.userGroup.deleteMany({ where: { userId: id } });
      if (groupIds.length > 0) {
        await this.prisma.userGroup.createMany({
          data: groupIds.map((gid) => ({ userId: id, groupId: gid })),
          skipDuplicates: true,
        });
      }
    }

    return await this.prisma.user.findUniqueOrThrow({
      where: { id },
      include: { groups: { include: { group: true } } },
    });
  }
}
