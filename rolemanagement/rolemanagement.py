import discord
from discord.ext import commands

class RoleManagement:
    """My custom cog that does stuff!"""

    def __init__(self, bot):
        self.bot = bot

    @commands.command(pass_context=True)
    async def togglerole(self, ctx, role : discord.Role):
        """Allows users to toggle a role on themselves.  You don't have to @mention the role, just type the name. Example: .togglerole lfg"""

        #Your code will go here
        try:
            if role in ctx.message.author.roles:
                await self.bot.remove_roles(ctx.message.author, role)
                await self.bot.say("Removed role "+str(role)+" from "+ctx.message.author.mention)
            else:
                await self.bot.add_roles(ctx.message.author, role)
                await self.bot.say("Added role "+str(role)+" to "+ctx.message.author.mention)
        except discord.Forbidden:
            await self.bot.say("I don't have permissions to manage that role.")

def setup(bot):
    bot.add_cog(RoleManagement(bot))
