// Usage:
//   HashPw <password>               → print a BCrypt hash (work factor 11)
//   HashPw verify <password> <hash> → print True/False whether they match
if (args.Length == 3 && args[0].Equals("verify", StringComparison.OrdinalIgnoreCase))
{
    Console.WriteLine(BCrypt.Net.BCrypt.Verify(args[1], args[2]));
    return 0;
}
if (args.Length == 1)
{
    Console.WriteLine(BCrypt.Net.BCrypt.HashPassword(args[0], workFactor: 11));
    return 0;
}
Console.Error.WriteLine("usage: HashPw <password> | HashPw verify <password> <hash>");
return 1;
