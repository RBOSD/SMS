import { PrismaClient, Role } from '@prisma/client';
import * as bcrypt from 'bcrypt';

const prisma = new PrismaClient();

async function main() {
    const username = 'admin';
    const newPassword = 'password123'; // Hardcoded recovery password
    const hashedPassword = await bcrypt.hash(newPassword, 10);

    console.log(`Resetting password for user: ${username}...`);

    try {
        const user = await prisma.user.upsert({
            where: { username },
            update: {
                passwordHash: hashedPassword,
                mustChangePassword: true,
                role: Role.MANAGER,
            },
            create: {
                username,
                passwordHash: hashedPassword,
                name: 'System Admin',
                role: Role.MANAGER,
                mustChangePassword: true,
            },
        });

        console.log(`Success! User '${username}' password set to '${newPassword}'`);
        console.log(`ID: ${user.id}`);
    } catch (e) {
        console.error('Error resetting password:', e);
    } finally {
        await prisma.$disconnect();
    }
}

main();
