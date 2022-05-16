import subprocess
import sys
import subprocess as sb

commit_id = sb.Popen(['git', 'merge-base' ,'FETCH_HEAD', 'HEAD'], stdout=sb.PIPE)
test=commit_id.communicate()[0]
print(test)
# sb.Popen(['git' , 'diff' ,'--name-status' ,test[:-1], 'HEAD'])
sb.Popen(['git' , 'diff' ,'--name-status' ,test.strip(), 'HEAD'])
args = sys.argv[1:]
print('args..',args)




subprocess.call(['java', '-jar', 'C://Users//UB217ZA//Downloads//excel-version-control//target//excel-version-control-1.0.jar', "ConvertToCSV"])

# output cell address and contents of cell
# def main():
#     args = sys.argv[1:]
#     print('args..',args)
#     print('args 1..',sys.argv[0:])
#     print('args length',len(args))
  

# if __name__ == '__main__':
#     main()