import { Injectable, OnModuleInit } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { Role } from '@prisma/client';
import * as bcrypt from 'bcrypt';
import { PrismaService } from './modules/prisma/prisma.service';

@Injectable()
export class AppInitService implements OnModuleInit {
  constructor(
    private readonly prisma: PrismaService,
    private readonly config: ConfigService,
  ) {}

  async onModuleInit() {
    // Ensure groups exist
    const adminGroupName = '系統管理群組';
    const defaultGroupName = '預設群組';

    // Ensure only one admin group (best-effort)
    await this.prisma.group.updateMany({
      where: { isAdminGroup: true, name: { not: adminGroupName } },
      data: { isAdminGroup: false },
    });

    const adminGroup = await this.prisma.group.upsert({
      where: { name: adminGroupName },
      update: { isAdminGroup: true },
      create: { name: adminGroupName, isAdminGroup: true },
    });

    const defaultGroup = await this.prisma.group.upsert({
      where: { name: defaultGroupName },
      update: {},
      create: { name: defaultGroupName, isAdminGroup: false },
    });

    const usersCount = await this.prisma.user.count();
    if (usersCount > 0) return;

    const username =
      this.config.get<string>('DEFAULT_ADMIN_USERNAME') || 'admin';
    const password =
      this.config.get<string>('DEFAULT_ADMIN_PASSWORD') || 'ChangeMe123';

    const passwordHash = await bcrypt.hash(password, 10);

    const adminUser = await this.prisma.user.create({
      data: {
        username,
        passwordHash,
        name: '系統管理員',
        role: Role.MANAGER,
        mustChangePassword: true,
      },
    });

    await this.prisma.userGroup.createMany({
      data: [
        { userId: adminUser.id, groupId: adminGroup.id },
        { userId: adminUser.id, groupId: defaultGroup.id },
      ],
      skipDuplicates: true,
    });
  }
}
