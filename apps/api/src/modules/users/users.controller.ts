import {
  Body,
  Controller,
  Get,
  Param,
  Post,
  Put,
  Req,
  UseGuards,
} from '@nestjs/common';
import { Role } from '@prisma/client';
import { AdminGuard } from '../../common/admin.guard';
import { getRequestIp } from '../../common/request-user';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AuditLogsService } from '../audit-logs/audit-logs.service';
import { CreateUserDto } from './dto/create-user.dto';
import { UpdateUserDto } from './dto/update-user.dto';
import { UsersService } from './users.service';

@Controller('admin/users')
@UseGuards(JwtAuthGuard, AdminGuard)
export class UsersAdminController {
  constructor(
    private readonly users: UsersService,
    private readonly audit: AuditLogsService,
  ) {}

  @Get()
  async list() {
    const rows = await this.users.list();
    return {
      data: rows.map((u) => {
        const groups = (u.groups || []).map((ug) => ug.group);
        const isAdmin = groups.some((g) => g.isAdminGroup === true);
        return {
          id: u.id,
          username: u.username,
          name: u.name,
          role: u.role,
          isAdmin,
          groupIds: groups.map((g) => g.id),
          groups: groups.map((g) => ({
            id: g.id,
            name: g.name,
            isAdminGroup: g.isAdminGroup,
          })),
          createdAt: u.createdAt,
          updatedAt: u.updatedAt,
        };
      }),
    };
  }

  @Post()
  async create(@Body() body: CreateUserDto, @Req() req: any) {
    const u = await this.users.create({
      username: body.username,
      password: body.password,
      name: body.name,
      role: body.role === 'MANAGER' ? Role.MANAGER : Role.VIEWER,
      groupIds: body.groupIds,
    });

    await this.audit.log({
      actorUserId: req.user?.userId,
      action: 'CREATE_USER',
      details: `create user ${u.username} (id=${u.id})`,
      ip: getRequestIp(req),
    });

    return { data: u };
  }

  @Put(':id')
  async update(
    @Param('id') idRaw: string,
    @Body() body: UpdateUserDto,
    @Req() req: any,
  ) {
    const id = parseInt(idRaw, 10);
    const u = await this.users.update(id, {
      name: body.name,
      password: body.password,
      role: body.role
        ? body.role === 'MANAGER'
          ? Role.MANAGER
          : Role.VIEWER
        : undefined,
      groupIds: body.groupIds,
    });

    await this.audit.log({
      actorUserId: req.user?.userId,
      action: 'UPDATE_USER',
      details: `update user ${u.username} (id=${u.id})`,
      ip: getRequestIp(req),
    });

    return { data: u };
  }
}
