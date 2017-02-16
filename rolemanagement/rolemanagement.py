import discord
from discord.ext import commands

class RoleManagement:
    """My custom cog that does stuff!"""

    def __init__(self, bot):
        self.bot = bot

    @commands.command()
    async def togglerole(self, ctx, role : discord.Role):
        """Allows users to toggle a role on themselves."""

        #Your code will go here
        already_has_role = role in ctx.message.author.roles
        await self.bot.say("Already in role? "+already_has_role)

def setup(bot):
    bot.add_cog(RoleManagement(bot))

