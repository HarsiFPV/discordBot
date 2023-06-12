"""Main file

N-word Counter bot
"""
import os
import datetime
import locale
import platform
import asyncio
import logging
from json import load
from pathlib import Path

import discord
import openpyxl
from discord.ext import commands

# Fetch bot token.
with Path("../config.json").open() as f:
    config = load(f)

TOKEN = config["DISCORD_TOKEN"]
USER_ID = 228954273218691074
excel_file = r"C:\Users\Tristan\PycharmProjects\pythonProject\N-Word-Counter-Bot\Connections.xlsx"

# Dictionnaire pour stocker les dates de connexion et de déconnexion par utilisateur
user_last_connection = {}
user_last_disconnection = {}

# Me and my alt account(s).
owner_ids = (354783154126716938, 691896247052927006)

intents = discord.Intents.default()
intents.members = True
intents.message_content = True
intents.presences = False

bot = commands.Bot(
    command_prefix=["nibba ", "n!"],
    case_insensitive=True,
    intents=intents,
    help_command=commands.MinimalHelpCommand()
)  # https://bit.ly/3rJiM2S

# Logging.
discord.utils.setup_logging(level=logging.INFO, root=False)
logger = logging.getLogger("discord")
logger.setLevel(logging.INFO)


async def main():
    async with bot:
        await load_extensions()
        await bot.start(TOKEN)


@bot.event
async def on_ready():
    """Display successful startup status"""
    logger.info(f"{bot.user.name} connected!")
    logger.info(f"Using Discord.py version {discord.__version__}")
    logger.info(f"Using Python version {platform.python_version()}")
    logger.info(f"Running on {platform.system()} {platform.release()} ({os.name})")


# Methode pour detecter la prise de parole d'un utilisateur
# @bot.event
# async def on_voice_state_update(member, before, after):
#    #logger.info(f"{member} speaked")
#    # Vérifie si le membre est celui à exclure
#    if member.id == 227041336396611584:
#        logger.info(f"{member.id} speaked")
#        # Vérifie si le membre a commencé à parler
#        if not after.self_mute and not after.afk:
#            try:
#                # Exclusion du membre
#                await asyncio.sleep(2)  # Pause de 1 seconde
#                channel = member.guild.get_channel(228954417204822016)  # Remplace CHANNEL_ID par l'ID du canal où envoyer le message
#                await channel.send("Tg David merde")
#               # await member.move_to(None, reason='Exclusion automatique (microphone)')
#               # print(f"L'utilisateur {member.name}#{member.discriminator} a été exclu (parole).")
#            except discord.Forbidden:
#                print(f"Je n'ai pas les permissions nécessaires pour exclure l'utilisateur {member.name}#{member.discriminator}.")

@bot.event
async def on_voice_state_update(member, before, after):
    if member.id == USER_ID:
        now = datetime.datetime.now()# Obtenir la date actuelle
        today = datetime.date.today()
        locale.setlocale(locale.LC_TIME, 'fr_FR')
        if before.channel is None and after.channel is not None:
            # Vérifier si une connexion a déjà été enregistrée pour la journée
            if member.id not in user_last_connection or user_last_connection[member.id] != today:
                formatted_date = now.strftime("%d %B %Y %H:%M")
                print(f'{member.name} a rejoint le salon {after.channel.name} {formatted_date}')
                print(user_last_connection)
                # Enregistrement dans le tableur Excel
                add_row_to_excel(excel_file, member.name, "Connexion", after.channel.name, formatted_date)
                user_last_connection[member.id] = today  # Mettre à jour la date de connexion pour l'utilisateur

        elif before.channel is not None and after.channel is None:
            # Vérifier si une déconnexion a déjà été enregistrée pour la journée
            if member.id not in user_last_disconnection or user_last_disconnection[member.id] != today:
                formatted_date = now.strftime("%d %B %Y %H:%M")
                print(f'{member.name} a quitté le salon {before.channel.name} {formatted_date}')
                print(user_last_disconnection)
                # Enregistrement dans le tableur Excel
                add_row_to_excel(excel_file, member.name, "Déconnexion", before.channel.name, formatted_date)
                user_last_disconnection[member.id] = today  # Mettre à jour la date de déconnexion pour l'utilisateur



def add_row_to_excel(file_path, member_name, action, channel_name, formatted_date):
    # Chargement du fichier Excel existant ou création d'un nouveau fichier
    try:
        workbook = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()

    # Sélection de la première feuille
    sheet = workbook.active

    # Création d'une nouvelle ligne avec les informations
    row = [member_name, action, channel_name, formatted_date]
    sheet.append(row)

    # Enregistrement du fichier Excel
    workbook.save(file_path)


@bot.event
async def on_command_error(ctx, error):
    logger.error(error)


@bot.command()
async def ping(ctx):
    """Pong back latency"""
    await ctx.send(f"_Pong!_ ({round(bot.latency * 1000, 1)} ms)")


@bot.command()
@commands.has_permissions(administrator=True)
async def load(context, extension):
    """(Bot dev only) Load a cog into the bot"""
    msg_success = f"File **load** of {extension}.py successful."
    msg_fail = "You do not have permission to do this"

    if context.author.id in owner_ids:
        await bot.load_extension(f"cogs.{extension}")
        logger.info(msg_success)
        await context.send(msg_success)
    else:
        await context.send(msg_fail)


@bot.command()
@commands.has_permissions(administrator=True)
async def unload(context, extension):
    """(Bot dev only) Unload a cog from the bot"""
    msg_success = f"File **unload** of {extension}.py successful."
    msg_fail = "You do not have permission to do this"

    if context.author.id in owner_ids:
        await bot.unload_extension(f"cogs.{extension}")
        logger.info(msg_success)
        await context.send(msg_success)
    else:
        await context.send(msg_fail)


@bot.command()
@commands.has_permissions(administrator=True)
async def reload(context, extension):
    """(Bot dev only) Reload a cog into the bot"""
    msg_success = f"File **reload** of {extension}.py successful."
    msg_fail = "You do not have permission to do this"

    if context.author.id in owner_ids:
        await bot.unload_extension(f"cogs.{extension}")
        await bot.load_extension(f"cogs.{extension}")
        logger.info(msg_success)
        await context.send(msg_success)
    else:
        await context.send(msg_fail)


# Load cogs into the bot.
async def load_extensions():
    for filename in os.listdir("./cogs"):
        if filename.endswith(".py"):
            await bot.load_extension(f"cogs.{filename[:-3]}")


asyncio.run(main())
