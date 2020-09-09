const { program } = require('commander');
const { createQueries } = require('./deployQueries');

program
  .version('0.0.1')
  .description('Deploy Query Builder');

  program
  .command('-createQueries <jiraTask> <date> <fileName>')
  .alias('cq')
  .description('Create Package')
  .action((tickets, date, fileName) => {
    createQueries(tickets, date, fileName);
  });

  program.parse(process.argv);
