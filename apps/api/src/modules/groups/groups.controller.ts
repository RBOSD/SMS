import {
  Body,
  Controller,
  Get,
  Param,
  Post,
  Put,
  UseGuards,
} from '@nestjs/common';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AdminGuard } from '../../common/admin.guard';
import { CreateGroupDto } from './dto/create-group.dto';
import { UpdateGroupDto } from './dto/update-group.dto';
import { GroupsService } from './groups.service';

@Controller('admin/groups')
@UseGuards(JwtAuthGuard, AdminGuard)
export class GroupsAdminController {
  constructor(private readonly groups: GroupsService) {}

  @Get()
  async list() {
    const rows = await this.groups.list();
    return {
      data: rows.map((g) => ({
        id: g.id,
        name: g.name,
        isAdminGroup: g.isAdminGroup,
        createdAt: g.createdAt,
        updatedAt: g.updatedAt,
      })),
    };
  }

  @Post()
  async create(@Body() body: CreateGroupDto) {
    const g = await this.groups.create({
      name: body.name,
      isAdminGroup: body.isAdminGroup,
    });
    return { data: g };
  }

  @Put(':id')
  async update(@Param('id') idRaw: string, @Body() body: UpdateGroupDto) {
    const id = parseInt(idRaw, 10);
    const g = await this.groups.update(id, {
      name: body.name,
      isAdminGroup: body.isAdminGroup,
    });
    return { data: g };
  }
}
