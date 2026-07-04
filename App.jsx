import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { db, auth, googleProvider } from './firebase';
import { ref, set, update, remove, onValue, off } from 'firebase/database';
import { onAuthStateChanged, signInWithPopup, signOut as fbSignOut } from 'firebase/auth';

// ─── CONFIG ──────────────────────────────────────────────────────────────────
const LOGO="data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAChAToDASIAAhEBAxEB/8QAHQABAAICAwEBAAAAAAAAAAAAAAYHBAgBAgUDCf/EAFUQAAEDAgEECQ4JCAgHAAAAAAACAwQBBQYHERITCBQhIjEyQWFxFRYjNkJRUlVigZSxwdE1cnN0gpGSobMXM0NEVoOT4SU0U2OEotLwRVR1o7LC8f/EABwBAQADAQEBAQEAAAAAAAAAAAAEBQYHAwgCAf/EADwRAAIBAgIFCQQJBQEBAAAAAAABAgMEBREGIUFRcRIxYYGRobHB0QcTIjUUFiMzNEJS4fAVMlNicqLx/9oADAMBAAIRAxEAPwDcsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHBGMY42w7hFyM1fJTjC5SVKa0GVOaVE5s/Fp5VCUFAbK/wCEcOfIyfW2edWbhFtFXjN5Oys5V6aWay5+lpE4/LRgHxnK9Dc9w/LRgHxnK9Dc9xqyCJ9JmYX64336Y9j9Tab8tGAfGcr0Nz3Ht4Rx7hzFUt6HZZTzzzLesXpx1t73Pm7qhp8XDsWe2+7fMKfiJP3TrylNJljhWk13eXcKM4xSe5PdxNi0nIBMN6AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABnABQGyv+EcOfIyfW2X+UBsr/hHDnyMn1tnjcfdsoNJ/llTq8UUiACvOSAuDYsdt91+Y0/EoU+XBsWO2+6/MafiUPSj94i60e+ZUuPkzYw5ODksjsAAAAAAAAAAAAAAAAAAAAAAAAABwo1ly4ZXZc+4vYcwtKXGt7KqtyZbSt/IVypRXkRz8vRwzLGxq3lTkU+t7iFfX9Kyp8up1LeXDi/KhgvDDy2LhdkPykU30aIirznnzbifpVK6ueyQgJ3LVheW8jw5MhDf3Uz+s11OprqGjtrBfHnJ9ngYuvpNd1H9mlFdpe1dkjeO4wvB9JX7jIh7JOR+t4TY/dTa+1BQIJTwOxf5O9+pF/r9/+vuXobTWfZCYMmV0LjDutsr4amqOI+tFc/3FjYcxXhzEbNXLHeYk3yW179PSmu7T6jRU+kV+REfRIiyHGX0cV1pdUqT0VoQa+jVvJfZScX2r+dZPt9J7iD+1ipLsfobu4gxT1JuG09o1e3qVaWno8PmMDr8b8Wr/AIv8iFJkyJlnsUyU+t59+0RnHXFcZSqp3a1ODg+K43f2t7VoxnqjJrmWw7Nh2GWlza06so65JPne0m3X434tX/F/kOvtvxav+N/IhIIH1lxH9fciY8Fs1+XvZaljvTVws67gtG1m0KUlWmvvcuc8qdjW2MK1cZl6Tz03tPvIMqdI6mt2+i9BhClOaPhVr3z4k6vpTc8iMaWWeWttbehcxFpYFRUnKfNnqXRxJn1+OeLf+9/Iy4GNre5XVy2XI3lcanvICdCFT0lxCMs3JPikSZ4JaNZKOXWXRGebkM0dZcQtC+KpJRGyu+EcO/IyfW2SbDV6kWiZ4cVf5xv205yLbKZxt2ZhtxtekhTD6kq5s7ZssOxaGI27aWUlzr06DnWm1hOzw+aetPLJ9aKUABIOKguDYsdt91+Y0/EoU+W1sZpUSHiO7SJT7EZnaHGcXRKfzlOWp6UfvEXWj3zKlx8mbJggF7ytYBtFKUexDHlL8GJne+9O4Qi87IyyN6bdnsM6Tub1T60s0+qmepf0sNu639lN+HidRrYpaUf76i7c/AvYGtyMsWUK53OG2xYoNpgvyGkKccZVn0aqpSu+crSleHvGyJ+bqyq2uSqZa9zzP3aX1K7z93nq3rIAAiEwAAAAAAAAAAAAAAAAAArvL/iFzDmTOe/FXq5UqqYbKuVNV8atOhFFGnRsnsv3F0w1YmqcRc1xSvM3XN6zWw3WjlJRtOXtbfdqOe6TVnO85GyKXedTsdQaAzxL8I5OsZYqidULPZlrjcVL7q0toXm4c2lXdPOxdhHEWFXkMX60vxNZ+bd3FIX0KTueY2wyM4gsd8wJaY9nfZ1kKE0y/HovfsLSmlK0rTp5eU65dcPT8UZPpNstUREmbr2nGkKXRPArdrnrwbmcyUcfrK793VilHPLc10tmwlo9Qlae8pSblln0PoNNTqWH+RfKP4ip6S37x+RfKP4ip6S37zQLELX/ACLtRnP6bd/432Fr2/tbw5/0WJ+GfQ+jkKXbLbZbfORqZUW1RmX08bRVROatM58z5bx9p4ncNfrl4n0fgyasKKf6V4AAFQWgB7uE8P8AVjTkSFrRFRvd7wqr3jIxbhpu1xNtwlrU1paLiVZs6c/BUs44PdStfpSj8PflvyK+eI0FX9xn8XcRoAFYWAIflve2zbcLaf6NEtv6lN5iYFT5T3JDmL9WtxzUoit6pKs+glVc+nm5O9n8xpdF4uVzUyeWUW+OuJgfaM1HBJavzR6tZFwAbM+dgYtzabcappyG0atClJSpClaXNTNSv+bMZRLsmGAm8fX163yLm5BYis651TSKKWvfUpmpn3KdO6WmC14W97CpUeSWfgTcPoTr3EacFm2QWFLtkdns9q2094TsmqWvsJpSv3khs2HcZYva2vZMN5ojmjvmIaGWemri+H66mzOEslGCMN0Q7Fs6JspH6zNrrl/VXcp5qE8SmiKUQihrLjSKHKfuYZ9Lfl+5v7bRubX288uhLz/Y10wXkAubc2HcsRXlhhcd1t5DEZFXFb2ufNVVc1KfebHAFBd3ta7kpVXnkaKzsKNnFxpLn5wACITAAAAAAAAAAAAAAAAAACldlvCXIwFb5dP1WenS+mhdPXmNYDebKFh1jFeELlYXq0RtprsavAcpXOhX10oaSXm3TLRcpNsuLC2ZMZ3VuNq8Kns5TbaN3MZUHR2p59TMJpPayjcKtsa70YQANIZcybbNmW+YiZb5ciFJRxXWF1StPnoWrhHL5iu16Ee8MR74x4SuxvfapuV+oqE7ES5sqFysqsUyXbX1xavOlJrw7DbzCeWvAt8ohp+eq0yq/o5qdCn8TifeWKw82+yh1lxC0L3UqQqikq85+fxIsHY1xPg+XRyyXJxlvuo7u+YX0o/+VM9daMxyzoSye5+ppLPSiWeVxHrXobM5Q+2RfyKPaeAZNxuTl7Ytd3fQhl+dbWHlNpz5kqWnPWlD43BLFnh7cv8APiWmL3KpK9+v4jdN2pwDELC5usUq0aMHKXKa1LM7TZ3lC3sKdWrJRjknrPiZMWDMmNLcYY3iOM4rMlCelVdwgF+ytWiHRbeGrN1Qf/525cTpSzSv/lUrnE+L8R4k+GLtIlI7ljitJ6EU3DZ4R7Mby4ynezVOO5a36IyOKe0Kzt84WseW9/MvU2Fcyr4QwbaVQKz+rM2i1K1VvzKQnPyVcrmR6yAXjLjf8QXONAYgxIFufkIbcaT2R1aaqpTNVVeD6NCnmUax5COxo0/CXRKfPXkLAsdhwRZ9qzL9iyl2m62mpgWdOlRC8+5rHV7mbPzHTI6P4fYWnuFFy+FpZ5vZ0autmC/r2IX9yqqkorNN8y29vYW0AD5jPoMFT5T5MhzF+13JDi2GIqFNN6e9SpWfSrSnPmp9RbBAMsEGJH6jXBtvQlStel9zd3yW6t6NPNnr9ZpdFpxjdTT2xa70YH2kRcsEk1slEr4AGzPnYFwbFntwuvzGn4lCny4Nix233X5jT8Sh6UfvEXWj3zKlx8mbGHJwclkdgAAAAAAAAAAAAAAAAAAAAAAAAOFEAyo5NLJjqJV1+m07m2nRZmtp33xVU7pJKbxfrJZ3G27rdoMBbm+bS+8lvSzcObOYXXtg/wDaizemo/1HvRlWpyVSlmn0Ea4jQqxdOrk105GpuOMmWLsIvOOzra5JhU4suHncazc/Kn6VCGG8vXpg/wDamzemo95DMQ4cyPYrlUbdk2VE99Wil2DMQ26tVeDcRXMuvxqVNRa6Q1Esrim+KXkZO70cpyedtUXBvzNSwX9iTY5SG+y4dxCh7+6mt6Kv4iP9JAZWTOdYHVuY2ucCwQW+Iui6PvP8zTaa51dKtEuaOLWlZfDPXu159hSVsHu6LynDVv1ZdpA223HHkNobcW4vepSnfKUrvUpykxZwPS1sol42uqMOsr3zUbQ1053oZpxKc68x2cxnDsbK4eBbZ1Npo6KrpJzOTneei+BqnMmnnIbIfckPLflSHH33N8p11dVKUrv1rU9ft63+i7/Rd/UeP2FH/d/+fV9y4lm3nK7MbhxrZhOD1MYix0RWpsnRclqbTTNTyE1+KVtPmS58xcydKkSn18Z99dXFq6a1M3DWHr3iS5bQsltfmP8Adavio51V4KULRTgPBGT1lu4ZRLtS53SqdJqzQv8A27qtOfe06SJGNnhz5NKPxvYlnJvp9WTGr3EUpVZfAtreUUuj9itsIYQxHiuXtex21+V4bvFaR0rruebhJnOwzk9wYytvFV3fxHe+L1Pta9W00ry3f918kwMZ5VL5e4nUe0NsYdsre9TBhb3ST3lKpw9Cc1Cvz2jTuK+uq+Qty5+t+naeUqtrb6qS5b3vm6l69h97i7HfmLciRdpsdyxrlOaP0q7tTvZWnHLxCbQ3puLkI0Up43GoWBk4yN4jxXVuZLbrZ7XX9O+jsrqf7tv2q3Ok2AseBcMYGw7LctMGm2tr1ouY9XTeXuZuHk6E5ivxTHbazpSjH4mk+ri//pPwvAbm7qxnL4VmufjsREQdDufLJ9GJZAhOWj4Nw5/i/W0TYhOWj4Nw5/i/W0X+jf4t/wDL8UYf2i/IqnGPiVqADcnziC4Nix233X5jT8ShT5cGxY7b7r8xp+JQ9KP3iLrR75lS4+TNjDk4OSyOwAAAAAAAAAAAAAAAAAAAAAAAAGvezCh01GHLhTuFvx1fSohVPVU14Nw9kNh9d/yZzaMN1XKgqTMaQnhVocan2KqNPaU8D6JvNHa6nZqO2La8/M55pJQcL1y2SSfkNEmGS2z3SXiqFeIsRHU61SmpEuW+ujceOlKqVrpuV3M+bk4TKi4VteG4jFzx04tDi06yNY4y9GU+mvBV2v6FH+Y8bFeLbpf2WYa9RDtbH9WtsNGrjtebuq86t0nzrSuU4Uubmb2dW/wIFOjG1anWeta1Fc/Xu8S5MpWX3NrLfghqi+5Vcn0b3923Xh6VfUUJdJ8+5zFz7jLfmSl8Z99dVKV56+oxiYZOsneIsbzP6Oj1i29tWi5Nfz6pPNTwq8yfuPKja2mG0+UtSXO3z9vkj0rXd3idXkvXuS5iJR2nJDyI7Dbj7ji9FKWkVUpSu9SlOEzpdivkNlcidZrjFZ7p1+MttP11obDvryfZEbdq2UUuuJHGu6za9eflryNN/wC90o3HuOcQYyuO2LxMrqW19giNVqllropy151bp/La+q3c+VThlT3vnfBH9urCjZw5NWedTctnFnex5Q8X2O29TLPeqQI39kxGaTw8ta6GetefhIzIfkSHlyH5Dj77i9Jx11dVKUrv1rXhOhZOAsmUi4Ms3fEy3LZa175thP8AWJnxaV4qPLV5j93l1Z4ZSlcVmoLa9/mzztbe8xKcbeknN7FnzehD8KYYveJ7jtS0RNdVG+ffVvWmE99aq7lKFz4SwZhzCmhIQ23fLuj9bfR2BhX903Xhr5aj3m1R49uRbLXEYt9vY/NxmOL0qrwqrzqOpxTST2h3N63Qsfgp7/zP08Truj+g1vYpVbv457ti9SR4WxO5EeebuK3323laWs41U19x98W4lj3CHSDbtOqVqzuKUnR4OTMRUGKWN3f0Z23K+F9vaa54Zb++VZLWuzsAAKkswQnLR8G4c/xfraJmQ/LUjQtuGqeHtpSftN09hoNG03dt/wCr8UYX2jNLA58Y+JWQANwfOQLg2LHbfdfmNPxKFPlwbFjtvuvzGn4lD0o/eIutHvmVLj5M2MOTg5LI7AAAAAAAAAAAAAAAAAAAAAAADEny40KE7MmvIjx2UVW444vRSlNOGtag/jeR93NDRXp5tDutI1Qxve8GYRxPc14DjtTLo46pSZrmZxm3K5UxkcFa58++7nuTIyz5X5eJ1PWTDjjkWycR13iuzP8ASjyeGvL4JUJscHwidOLnXeWf5fX07dxicaxqFSSp0Enl+b09ew+suRImS3pcp9x999ek666uqlKVXlrWvCdG0LcdQ2htxbji9FKE75SlV4KUoehhyx3TEl3YtFniuSpb/FSnipTyqVXkpzm0eTDJbY8BQ63a6vszLmhGk5MczJajJzbtG8/B8bh6C0v8To2MMueWxL+aipw/C61/LPmjtZBMk2QxyRqrvjZtbLXGatunvlc7tacHxPtd49HKnligWCEvDGAm4+vYRqVSmkUozG8lqlNyqufi05yMZacsUi/66wYXdcjWimk2/JTvXJPNTwW/vqUyQrbD615NXF91R2Lj/OJOusRo2UHb2PXLa+H84H3mSX5kt6XKfceffXpOOOrqpald+ta8JxFYkS5iI8WO4884vRbaaRpKWqvJSh3tkKZc7kxb7fHckyX16tppvhVU2wyNZK4GC4qbhcdXMvryN+7TiMJ8Bv215SdiOI0rGnr1vYiBhuG1b+puW1/zaV3gjJxDwrRm4YliNzb1VKXGISt8zD7ynORxfNwUJZKfkSHVyH3HFrX3Sj3MonbIv5JHtPAPnHSrGbvEr6aryzUW0lsR3rR3CbWwtIqjHJta3tYABmS/Zl1tzlbO3dEb9GtU255PeqYhYWT5tt3Da21o0kLecopKvMebe8Fro4p21uUU3/YOdz0VNHXwCrO2p3Fus+VFZrbn0ehTUsWhCvOjWeWTeT2EPB6LmH723/wx/wCiZUDCl3mO9kY2qjwnPcVEMOu5y5Mabz4MsJXtvBcpzWXFHlW2HIuExuGwjfr/AMvfrUj+ydiohvYZiM8Rth9P3tl3WGxw7QzXU795dN+6qm+Vzc1Co9k5bbncLjYNoW2XMohp/S2uypzR3W82fNTcN7hGDuwoSlU1zl3Lcc104vne2E4U0+Sssul5rWUOD1+tvEf7PXX0Jz3DrbxH+z119Cc9xYcmW44r9FrfofYzyC4Niz24XX5jT8ShW3W3iP8AZ66+hOe4tTY1Wu6W/FFzcnW2dFbXD0ULkRlNpUrTpuZ60PSjGXLWaLjALerDEaTlFrXue5mwJycJOSxOtAAAAAAAAAAAAAAAAAAAAAGFdbhDtdueuFxkNxorCdN1xxWilCacpqZloypTMbTnIFuccjWJhfYmuKqTWndueynJ0m28yLGmM6iSw0+zXjNuIopNfNUwet3D3iO2+io9xY4deUbSfvJw5UtmvmKzErOteQ93CfJW3VzmhWk2SPAOELvjK+otdnb8p+QqnYmEd9VfVTlN0ut3D/iK2eiI9xlQbfAgUUiFCjxkr41GWqJz9OYuauk7lBqnTye/Mo6OiqjNOpUzXQsvMjGCcKYcydYac1C22UNtaybPfzJU5mpuqXXkpzchr1lryrSMZTF2e0OORrE2rg4qpKqd0vyO9T6+bbCXFjS2FMSmG32V03zbqaKSrzVMLrdw94itvoiPcVFlfwo1XWrRc5cS5vcPqVqKoUZKEeH8/c0J0j6sNuSHkR2G9c+4tKUNp3ylqrwUpQ3w63cP+Irb6Ij3HZmxWOO6h9i025laN1K0RkJUjnpXMXb0ojlqpd/7FFHROWeur3fuQLIdkzYwbadv3FtC79KT2VW4qkZNf0aPbXlLSAMvXr1Lio6lR5tmstrenbU1TprJIrHKL2yL+RR7TwD38ovbIv5FHtPAOO4z+Pq/9M6Nhv4WnwQABWkxli5OO1z9+r2EoIvk47XP36vYSg65g/4Cj/yjAYh+KqcWAAWRDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArwAAFY5Re2RfyKPaeAW3Ns9rmPa+VCYec4NJVD5dbtk8WR/smJvtGK1zczqxmkpPPaaO1xynQoxpuL1LoKpBa3W7ZPFkf7I63bJ4sj/ZIv1QuP8ke8kfWKl+h9x52Tvtfr8uv2EoMWHDjw2dVGYQ0jPn0UmUbWxoO3toUm83FJGbuaqrVpVFtYABLPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//2Q==";
const NV="#1e3a5f",AM="#f59e0b",GN="#16a34a",RD="#dc2626",PU="#6d28d9";
const BUILD_TAG="v4 \u00b7 user activity log";
const QUOTES=[
  {q:"Great works are performed not by strength, but by perseverance.",a:"Samuel Johnson"},
  {q:"The only way to do great work is to love what you do.",a:"Steve Jobs"},
  {q:"Every brick laid today is progress built forever.",a:""},
  {q:"Infrastructure is the foundation on which a nation's dreams are built.",a:""},
  {q:"Hard work is the bridge between goals and achievement.",a:""},
  {q:"Progress is the sum of small efforts repeated day in and day out.",a:"Robert Collier"},
  {q:"The road to success is always under construction.",a:""},
  {q:"Build today with precision — tomorrow will thank you.",a:""},
  {q:"Excellence is not a destination but a continuous journey that never ends.",a:"Brian Tracy"},
  {q:"The difference between ordinary and extraordinary is that little extra.",a:"Jimmy Johnson"},
  {q:"Discipline is the bridge between goals and accomplishment.",a:"Jim Rohn"},
  {q:"Quality means doing it right when no one is looking.",a:"Henry Ford"},
  {q:"Do not wait; the time will never be 'just right.' Start where you stand.",a:"Napoleon Hill"},
  {q:"Success is the result of perfection, hard work, learning from failure, and persistence.",a:"Colin Powell"},
  {q:"The strength of a structure lies in the dedication of those who build it.",a:""},
  {q:"Challenges are what make life interesting. Overcoming them is what makes it meaningful.",a:"Joshua Marine"},
  {q:"It always seems impossible until it is done.",a:"Nelson Mandela"},
  {q:"Your work is a reflection of your character. Make it count.",a:""},
  {q:"A man who dares to waste one hour of time has not discovered the value of life.",a:"Charles Darwin"},
  {q:"The secret of getting ahead is getting started.",a:"Mark Twain"},
  {q:"We build too many walls and not enough bridges.",a:"Isaac Newton"},
  {q:"Engineering is not only the study of why things work, but why things fail — and how to fix them.",a:""},
  {q:"Safety is not a gadget but a state of mind.",a:"Eleanor Everet"},
  {q:"Teamwork makes the dream work. Every hand on site matters.",a:""},
  {q:"Plan your work for today and every day, then work your plan.",a:"Margaret Thatcher"},
  {q:"The man who moves a mountain begins by carrying away small stones.",a:"Confucius"},
  {q:"Strive not to be a success, but rather to be of value.",a:"Albert Einstein"},
  {q:"There are no shortcuts to any place worth going.",a:"Beverly Sills"},
  {q:"Good things come to people who wait, but better things come to those who go out and get them.",a:""},
  {q:"Every expert was once a beginner. Every pro started as an amateur.",a:""},
  {q:"The future belongs to those who believe in the beauty of their dreams.",a:"Eleanor Roosevelt"},
];
const RTDB_URL="https://spl-dpr-default-rtdb.asia-southeast1.firebasedatabase.app";

// ── Auth-token helpers ───────────────────────────────────────────────────────
// Database rules require a signed-in Google user (auth != null). Every REST
// call must carry the current user's Firebase ID token as ?auth=<token>.
async function idToken(){
  try{ return auth.currentUser ? await auth.currentUser.getIdToken() : ''; }
  catch(e){ return ''; }
}
// Append ?auth / &auth to a full RTDB REST URL
export async function authedUrl(url){
  const t=await idToken();
  if(!t) return url;
  return url + (url.includes('?') ? '&' : '?') + 'auth=' + t;
}

// REST API write helpers — authenticated with the Google ID token
// Works on every device, every network, every browser
async function rtdbPut(path, data){
  const r=await fetch(await authedUrl(RTDB_URL+'/'+path+'.json'),{
    method:'PUT',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify(data)
  });
  if(!r.ok){const t=await r.text();throw new Error('RTDB PUT '+r.status+': '+t);}
  return r.json();
}
async function rtdbPatch(path, data){
  const r=await fetch(await authedUrl(RTDB_URL+'/'+path+'.json'),{
    method:'PATCH',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify(data)
  });
  if(!r.ok){const t=await r.text();throw new Error('RTDB PATCH '+r.status+': '+t);}
  return r.json();
}
async function rtdbDelete(path){
  const r=await fetch(await authedUrl(RTDB_URL+'/'+path+'.json'),{method:'DELETE'});
  if(!r.ok) throw new Error('RTDB DELETE '+r.status);
}
// Fetch every project's submissions in one shot (for cross-project screens:
// Analytics, Performance, Reports). Tags each submission with its project.
async function fetchAllSubmissions(projects){
  const chunks=await Promise.all((projects||[]).map(async p=>{
    try{
      const r=await fetch(await authedUrl(RTDB_URL+'/projects/'+p.id+'/submissions.json'));
      if(!r.ok) throw new Error('HTTP '+r.status);
      const obj=await r.json()||{};
      return Object.values(obj).filter(Boolean).map(s=>({...s,_projectId:p.id,_projectName:p.name}));
    }catch(e){ return []; }
  }));
  return chunks.flat();
}

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const ASSET_GROUPS=["Excavator","Motor Grader","Soil Compactor","Crane Hydra 14 MT","Back Hoe Loader","Bitumen Sprayer","Concrete Boom Pump","Crane F23","Front End Wheel Loader","Generator","Hydraulic Piling Rig","Mini Tandem Roller","Pneumatic Tyre Roller","Tipper 10 Wheeler","Tipper 12 Wheeler","Tipper 6 Wheeler","Sensor Paver 6M","Sensor Paver 7M","Sensor Paver 7.5M","Sensor Paver 9M","Tandem Roller","Transit Mixer","Water Tanker"];
const LABOUR_TYPES=["Helper","Fitter","Mazdoor Skilled","Mazdoor Unskilled","Other"];
const VEHICLE_TYPES=["Tipper 6W","Tipper 10W","Tipper 12W","Transit Mixer","Water Tanker","Trailer","Dumper","Other"];
const SIDES=["RHS","LHS","BHS","Median","RE","N/A"];
const UNITS=["Cum","Sqm","Rm","Nos","MT","Kg","Rmt"];
const DEPTS=["HW - Highway","STR - Structures","Laying","Base Camp / Precast Yard"];
const DESIGNATIONS=["Senior Engineer","Site Engineer","Junior Engineer","Assistant Engineer","Supervisor","Resident Engineer"];
const INCHARGE_OPTS=["Rajesh Kumar","K. Subramaniam","Venkatesh","Anand","Other"];
const PROD_MATS=[{n:"M15 RMC (FA)",u:"Cum"},{n:"M20 (Kerb) RMC (FA)",u:"Cum"},{n:"M25 RMC (FA)",u:"Cum"},{n:"M30 RMC (FA)",u:"Cum"},{n:"M35 RMC (FA)",u:"Cum"},{n:"M35(Pile) RMC (FA)",u:"Cum"},{n:"M40 RMC (FA)",u:"Cum"},{n:"M40(Pile) RMC (FA)",u:"Cum"},{n:"M50 RMC (FA)",u:"Cum"},{n:"BC GR-II CRMB60",u:"MT"},{n:"DBM GR-II VG40",u:"MT"},{n:"WMM",u:"MT"}];
const BULK_MATS=["Aggregate","Gravel","GSB","Over Burden (OB)","Earth / Soil","WBM Material","M-Sand","Concrete Blocks","Steel","CRS Steel","HT Strand","MS Liner Plate","Other"];
const DEFAULT_WT=["C&G","BT Dismantling","WBM Dismantling","Excavation","Embankment","Construction of Embk from Excavation","Culvert Back Filling-Dust","RE Wall Filling","Median Filling","Slope protection work","Geo Cell Fixing","Coirmat","Vetiver","GSB","WMM","Prime Coat Over Granular Surface","Tack Coat Over Granular Surface","Tack Coat Over Bitumen Surface","DBM","Patch Work","BC","RE Wall Levelling pad & Surface Drain","PCC","RE Panel Casting","RE Panel Erection","Perforated Pipe","Friction Slab With Crash Barrier","Crash Barrier","Friction Slab PCC","Friction Slab Raft","Friction Slab Haunch","Coping Beam","Soil Excavation For Box & V-Drain","Box Drain PCC","Box Drain Raft Concrete","Box Drain Wall Concrete","Box Drain Top Slab Concrete","V - Drain","Kerb","Kerb Rectification","Kerb Precast","Plantation","Diversion","Diversion - Soil","Diversion - GSB","Diversion - WMM","Diversion - DBM","Diversion - BC","Slope protection work (Emb)","Paver Block Erection","MBCB","Misc Works","Sub-Structures","Superstructure","Approach Slab PCC","Approach Slab Crash Barrier","Culvert Excavation","Culvert PCC","Culvert Raft Concrete","Culvert Wall Concrete","Culvert Slab Concrete","Flexible Apron","Parapet","Culvert Foundation","Culvert Sub structure","Culvert Super Structure","MNB Foundation","MNB SubStructure","MNB Super Structure","MNB Girder Erection","MJB PCC","MJB Foundation","MJB Sub structure","MJB Super structure","MJB Girder Casting","Pile Work","MJB Crash Barrier","Geo Composite","Sacrificial Slab Casting","Miscellaneous","Girder Erection","Culvert Box Segment Erection","MJB Work Progress","PVD Installation","Sub Grade","MNB-Backfilling","Cable Erection","HDPE Laying","Primer 1st Coating","Surface Drain Erection","Surface Drain Casting","Light pole Casting","Light Pole Erection","MNB PCC","Back Filling","PVD Filling","Utility"];
const ROLE_CAPS={engineer:{fill:true,approve:false,download:false,manage:false,settings:false,users:false},incharge:{fill:true,approve:true,download:false,manage:false,settings:false,users:false},management:{fill:false,approve:true,download:true,manage:true,settings:false,users:false},admin:{fill:true,approve:true,download:true,manage:true,settings:true,users:true}};
const ROLE_LABELS={engineer:"Engineer",incharge:"Incharge",management:"Management",admin:"Admin"};
const roleKey=r=>(r||"").trim().toLowerCase();
const isSuperAdminUser=u=>!!u&&(u.id==="superadmin"||(roleKey(u.role)==="admin"&&(u.name||"").trim().toLowerCase()==="super admin"));
const fmtTs=ts=>{if(!ts)return"";const d=new Date(ts);if(isNaN(d.getTime()))return"";return d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"})+", "+d.toLocaleTimeString("en-US",{hour:"numeric",minute:"2-digit",hour12:true}).toLowerCase();};
const WEATHER_OPTS=["☀️ Clear","🌤️ Partly Cloudy","☁️ Overcast","🌦️ Light Rain","🌧️ Heavy Rain","🌫️ Fog","🌡️ Extreme Heat"];
const PROBLEM_OPTS=[
  "✅ No Issues",
  "🚛 Vehicle Breakdown — machine stopped mid-day",
  "🪨 Material Shortage — material not available at site",
  "🚧 Workfront Not Available — location blocked or occupied",
  "⚡ Power Failure — generator or electrical issue",
  "🌧️ Rain Stoppage — work halted due to rain",
  "⚙️ Equipment Breakdown — plant/machinery failure",
  "👷 Labour Shortage — insufficient workforce",
  "🚦 Road Block — access to work zone blocked",
  "📐 Survey Delay — surveyor or layout not ready",
  "📋 Permission / Approval Pending",
  "🏗️ Third Party Interference",
  "📝 Other — mentioned in remarks",
];
const DEFAULT_ENG=[{id:"e1",name:"Viji",dept:"HW - Highway",incharge:"Rajesh Kumar",designation:"Site Engineer"},{id:"e2",name:"Mohan",dept:"HW - Highway",incharge:"Rajesh Kumar",designation:"Site Engineer"},{id:"e3",name:"Jayakumar",dept:"HW - Highway",incharge:"Rajesh Kumar",designation:"Site Engineer"},{id:"e4",name:"Kalaiyarasan",dept:"HW - Highway",incharge:"K. Subramaniam",designation:"Site Engineer"},{id:"e5",name:"Hareesh",dept:"HW - Highway",incharge:"K. Subramaniam",designation:"Site Engineer"},{id:"e6",name:"Gunasekar",dept:"HW - Highway",incharge:"K. Subramaniam",designation:"Site Engineer"},{id:"e7",name:"Madhan",dept:"HW - Highway",incharge:"Rajesh Kumar",designation:"Site Engineer"},{id:"e8",name:"Silambarasan",dept:"HW - Highway",incharge:"Rajesh Kumar",designation:"Junior Engineer"},{id:"e9",name:"Bhujanga",dept:"HW - Highway",incharge:"Venkatesh",designation:"Site Engineer"},{id:"e10",name:"Sivagandan",dept:"HW - Highway",incharge:"Venkatesh",designation:"Site Engineer"},{id:"e11",name:"Pawan Kumar",dept:"HW - Highway",incharge:"K. Subramaniam",designation:"Junior Engineer"},{id:"e12",name:"Ramesh",dept:"HW - Highway",incharge:"Rajesh Kumar",designation:"Site Engineer"},{id:"e13",name:"Murugesh Pandi",dept:"HW - Highway",incharge:"K. Subramaniam",designation:"Site Engineer"},{id:"e14",name:"Subash",dept:"HW - Highway",incharge:"K. Subramaniam",designation:"Site Engineer"},{id:"e15",name:"Mareeshwaran",dept:"HW - Highway",incharge:"K. Subramaniam",designation:"Site Engineer"},{id:"e16",name:"Vignesh",dept:"HW - Highway",incharge:"Rajesh Kumar",designation:"Junior Engineer"},{id:"e17",name:"Robert",dept:"HW - Highway",incharge:"K. Subramaniam",designation:"Site Engineer"},{id:"e18",name:"Balaji",dept:"STR - Structures",incharge:"Venkatesh",designation:"Senior Engineer"},{id:"e19",name:"Subodh",dept:"STR - Structures",incharge:"Venkatesh",designation:"Site Engineer"},{id:"e20",name:"Ramprakash",dept:"STR - Structures",incharge:"Venkatesh",designation:"Site Engineer"},{id:"e21",name:"Ganesh Pradhan",dept:"STR - Structures",incharge:"Venkatesh",designation:"Site Engineer"},{id:"e22",name:"Manikandan",dept:"STR - Structures",incharge:"K. Subramaniam",designation:"Site Engineer"},{id:"e23",name:"Sathyan",dept:"STR - Structures",incharge:"Venkatesh",designation:"Junior Engineer"},{id:"e24",name:"Vishnu",dept:"STR - Structures",incharge:"Venkatesh",designation:"Junior Engineer"},{id:"e25",name:"Guna",dept:"STR - Structures",incharge:"K. Subramaniam",designation:"Junior Engineer"},{id:"e26",name:"Ajith",dept:"Base Camp / Precast Yard",incharge:"Venkatesh",designation:"Site Engineer"},{id:"e27",name:"Surya",dept:"Base Camp / Precast Yard",incharge:"Rajesh Kumar",designation:"Site Engineer"}];
const DEFAULT_USERS=[{id:"u1",name:"Rajesh Kumar",role:"incharge",pin:"1111"},{id:"u2",name:"K. Subramaniam",role:"incharge",pin:"2222"},{id:"u3",name:"Venkatesh",role:"incharge",pin:"3333"},{id:"u4",name:"Anand",role:"management",pin:"4444"},{id:"u5",name:"Admin",role:"admin",pin:"0000"}];

const uid=()=>Math.random().toString(36).slice(2,9);

// Firebase SDK onValue() converts arrays → objects {"0":…,"1":…}
// toArr converts both back to a real JS array safely
const toArr=v=>!v?[]:Array.isArray(v)?v:Object.values(v);
const normalizeAct=a=>({...a,assets:toArr(a.assets),contractors:toArr(a.contractors).map(c=>({...c,entries:toArr(c.entries)}))});
const normalizeSub=s=>({...s,activities:toArr(s.activities).map(normalizeAct),matTxs:toArr(s.matTxs).map(m=>({...m,assets:toArr(m.assets),contractors:toArr(m.contractors).map(c=>({...c,entries:toArr(c.entries)}))}))});

const mkAsset=()=>({_k:uid(),group:"",name:"",unit:"Hour",hours:"",fuel:""});
const mkAct=()=>({_k:uid(),actType:"",actCustom:"",desc:"",chFrom:"",chTo:"",side:"RHS",cw:"",unit:"Cum",nos:"1",length:"",width:"",depth:"",theoQty:"",prodMat:"",prodMatCustom:"",prodQty:"",remarks:"",assets:[],contractors:[]});
const mkMatTx=()=>({_k:uid(),txType:"receive",material:"",matCustom:"",recvCH:"",source:"",recvQty:"",recvUnit:"Cum",recvTransporter:"",recvVehicle:"",recvLoads:"",sendFromCH:"",sendToCH:"",sendQty:"",sendUnit:"Cum",sendTransporter:"",sendVehicle:"",sendLoads:"",assets:[],contractors:[],remarks:""});
const mkHdr=()=>({engineer:"",date:new Date().toISOString().slice(0,10),shift:"Day",dept:"",incharge:"",engCustom:"",weather:"☀️ Clear",difficulty:""});

// ─── MOBILE HOOK ──────────────────────────────────────────────────────────────
function useMobile(){
  const [m,setM]=useState(window.innerWidth<768);
  useEffect(()=>{const h=()=>setM(window.innerWidth<768);window.addEventListener("resize",h);return()=>window.removeEventListener("resize",h);},[]);
  return m;
}

// ─── UI ATOMS ─────────────────────────────────────────────────────────────────
function L({t,req}){return <label style={{fontSize:"13px",color:"#374151",fontWeight:"600",display:"block",marginBottom:"5px"}}>{t}{req&&<span style={{color:RD,marginLeft:"2px"}}>*</span>}</label>;}
function Inp({style,...p}){return <input inputMode={p.type==="number"?"decimal":p.type==="tel"?"tel":undefined} {...p} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px",color:"#111827",background:"#fff",...(style||{})}}/>;}
function Sel({val,onChange,opts=[],placeholder,children}){
  return <select value={val} onChange={e=>onChange(e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px",color:val?"#111827":"#6b7280",background:"#fff",minWidth:0,maxWidth:"100%",WebkitAppearance:"none",MozAppearance:"none",appearance:"none",backgroundImage:"url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%236b7280' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E\")",backgroundRepeat:"no-repeat",backgroundPosition:"right 10px center",paddingRight:"34px",overflow:"hidden",textOverflow:"ellipsis"}}>
    {placeholder&&<option value="">{placeholder}</option>}
    {children||opts.map(o=><option key={o}>{o}</option>)}
  </select>;
}
function Seg({opts,val,onChange,small}){return <div style={{display:"flex",borderRadius:"8px",overflow:"hidden",border:"1.5px solid #d1d5db"}}>{opts.map(o=><button key={o} onClick={()=>onChange(o)} style={{flex:1,padding:small?"8px 4px":"11px 4px",border:"none",fontSize:small?"12px":"14px",cursor:"pointer",fontWeight:val===o?"700":"400",background:val===o?NV:"#fff",color:val===o?"#fff":"#6b7280"}}>{o}</button>)}</div>;}
function F({lbl,req,col,children}){return <div style={{gridColumn:col?`span ${col}`:undefined}}>{lbl&&<L t={lbl} req={req}/>}{children}</div>;}
function Grid({cols,children}){return <div style={{display:"grid",gridTemplateColumns:cols||"1fr",gap:"14px"}}>{children}</div>;}
function initials(n){return(n||"?").split(" ").map(x=>x[0]||"").join("").slice(0,2).toUpperCase();}
function Av({name,sz}){const s=sz||32;return <div style={{width:s,height:s,borderRadius:"50%",background:"#fef3c7",display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.round(s*.34)+"px",fontWeight:"700",color:"#d97706",flexShrink:0}}>{initials(name)}</div>;}
function DeptB({dept}){if(!dept)return null;const d=dept.toLowerCase();const[bg,fg]=d.includes("hw")?["#dbeafe","#1e40af"]:d.includes("str")?["#ede9fe","#5b21b6"]:d.includes("lay")?["#dcfce7","#166534"]:["#fef9c3","#854d0e"];const s=d.includes("hw")?"HW":d.includes("str")?"STR":d.includes("lay")?"Laying":"BaseCamp";return <span style={{background:bg,color:fg,padding:"3px 8px",borderRadius:"5px",fontSize:"11px",fontWeight:"700"}}>{s}</span>;}
function RoleB({role}){const m={engineer:["#dbeafe","#1e40af"],incharge:["#dcfce7","#166534"],management:["#ede9fe","#5b21b6"],admin:["#fef9c3","#854d0e"]};const[bg,fg]=m[roleKey(role)]||["#f3f4f6","#374151"];return <span style={{background:bg,color:fg,padding:"3px 8px",borderRadius:"5px",fontSize:"11px",fontWeight:"700"}}>{ROLE_LABELS[roleKey(role)]||role}</span>;}
function Pill({label,color,bg}){return <span style={{background:bg||"#f3f4f6",color:color||"#374151",padding:"3px 9px",borderRadius:"5px",fontSize:"11px",fontWeight:"700",whiteSpace:"nowrap"}}>{label}</span>;}
function Card({children,style}){return <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"16px",...(style||{})}}>{children}</div>;}
function SecHead({title,subtitle,color}){return <div style={{fontWeight:"700",fontSize:"13px",color:color||NV,marginBottom:"14px",paddingBottom:"8px",borderBottom:`2px solid ${color||NV}20`}}>{title}{subtitle&&<div style={{fontWeight:"400",fontSize:"11px",color:"#6b7280",marginTop:"2px"}}>{subtitle}</div>}</div>;}

// ─── SECTION BOX (colored header, no overflow:hidden) ─────────────────────────
function SB({emoji,title,sub,color,children}){
  const c=color||NV;
  return(
    <div style={{border:`2px solid ${c}25`,borderRadius:"12px",marginBottom:"18px"}}>
      <div style={{background:c,padding:"12px 16px",borderRadius:"10px 10px 0 0",display:"flex",alignItems:"center",gap:"10px"}}>
        <span style={{fontSize:"22px",lineHeight:1}}>{emoji}</span>
        <div><div style={{fontWeight:"700",fontSize:"14px",color:"#fff"}}>{title}</div>{sub&&<div style={{fontSize:"11px",color:"rgba(255,255,255,.75)",marginTop:"2px"}}>{sub}</div>}</div>
      </div>
      <div style={{padding:"16px",background:"#fff"}}>{children}</div>
    </div>
  );
}

// ─── WORK TYPE AUTOCOMPLETE (fixed position dropdown) ─────────────────────────
function WTInput({value,onChange,wt}){
  const [q,setQ]=useState(value||"");
  const [open,setOpen]=useState(false);
  const [pos,setPos]=useState({top:0,left:0,width:300});
  const ref=useRef(null);
  useEffect(()=>{setQ(value||"");},[value]);
  const filtered=q.length>0?wt.filter(t=>t.toLowerCase().includes(q.toLowerCase())).slice(0,16):wt.slice(0,16);
  function show(){if(ref.current){const r=ref.current.getBoundingClientRect();setPos({top:r.bottom+4,left:r.left,width:r.width});}setOpen(true);}
  function pick(v){setQ(v);onChange(v);setOpen(false);}
  return(
    <div>
      <input ref={ref} value={q} onChange={e=>{setQ(e.target.value);onChange(e.target.value);show();}} onFocus={show} onBlur={()=>setTimeout(()=>setOpen(false),200)}
        placeholder="Type to search — e.g. GSB, Pile, Embankment…"
        style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"2px solid "+PU,fontSize:"15px",fontWeight:"600",color:"#111827",background:"#faf5ff"}}/>
      {open&&filtered.length>0&&(
        <div style={{position:"fixed",top:pos.top,left:pos.left,width:pos.width,background:"#fff",border:"2px solid "+PU,borderRadius:"10px",zIndex:99999,maxHeight:"260px",overflowY:"auto",boxShadow:"0 12px 40px rgba(0,0,0,.2)"}}>
          {filtered.map(t=>{const i=t.toLowerCase().indexOf(q.toLowerCase());return(
            <div key={t} onMouseDown={()=>pick(t)} style={{padding:"12px 14px",cursor:"pointer",fontSize:"14px",color:"#111827",borderBottom:"1px solid #f3f4f6",background:"#fff"}}>
              {q&&i>=0?(<>{t.slice(0,i)}<span style={{background:"#fef08a",fontWeight:"800",borderRadius:"2px"}}>{t.slice(i,i+q.length)}</span>{t.slice(i+q.length)}</>):t}
            </div>);
          })}
        </div>
      )}
    </div>
  );
}

// ─── ASSET SECTION ────────────────────────────────────────────────────────────
function AssetSection({assets,onChange,assetGroups}){
  const [show,setShow]=useState(false);
  const [d,setD]=useState(mkAsset());
  const mobile=useMobile();
  const grps=assetGroups||ASSET_GROUPS;
  function add(){if(!d.group&&!d.name)return;onChange([...assets,{...d,_k:uid()}]);setD(mkAsset());setShow(false);}
  return(
    <div>
      {assets.map((a,i)=>(
        <div key={a._k} style={{display:"flex",gap:"10px",alignItems:"flex-start",padding:"10px 14px",borderRadius:"8px",background:"#fffbeb",border:"1px solid #fde68a",marginBottom:"8px"}}>
          <span style={{fontWeight:"700",color:"#d97706",minWidth:"20px",fontSize:"13px",paddingTop:"1px"}}>{i+1}.</span>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:"700",fontSize:"14px",color:"#92400e",wordBreak:"break-word"}}>{a.group||"—"}</div>
            {a.name&&<div style={{color:"#6b7280",fontSize:"12px",marginTop:"2px",wordBreak:"break-word"}}>ID: {a.name}</div>}
            <div style={{display:"flex",gap:"8px",marginTop:"5px",flexWrap:"wrap"}}>
              {a.hours&&<span style={{background:"#fff",border:"1px solid #fde68a",padding:"3px 9px",borderRadius:"6px",fontSize:"12px",fontWeight:"600",color:"#92400e"}}>{a.hours} {a.unit||"Hour"}s</span>}
              {a.fuel&&<span style={{background:"#fef2f2",border:"1px solid #fca5a5",padding:"3px 9px",borderRadius:"6px",fontSize:"12px",color:RD,fontWeight:"600"}}>⛽ {a.fuel}L</span>}
            </div>
          </div>
          <button onClick={()=>onChange(assets.filter(x=>x._k!==a._k))} style={{padding:"6px 10px",borderRadius:"6px",border:"1px solid #fca5a5",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"13px",flexShrink:0}}>✕</button>
        </div>
      ))}
      {show?(
        <div style={{background:"#f8fafc",border:"1px solid #d1d5db",borderRadius:"10px",padding:"14px",marginTop:"8px"}}>
          <Grid cols={mobile?"1fr":"1fr 1fr"}>
            <F lbl="Asset Group">
              <select value={d.group} onChange={e=>setD(p=>({...p,group:e.target.value}))} style={{width:"100%",boxSizing:"border-box",padding:"13px 36px 13px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"16px",color:d.group?"#111827":"#6b7280",background:"#fff",WebkitAppearance:"none",MozAppearance:"none",appearance:"none",backgroundImage:"url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%236b7280' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E\")",backgroundRepeat:"no-repeat",backgroundPosition:"right 12px center",minHeight:"50px"}}>
                <option value="">— Select Group —</option>
                {grps.map(o=><option key={o} value={o}>{o}</option>)}
              </select>
            </F>
            <F lbl="Asset Name / ID (e.g. POC-056)"><Inp value={d.name} onChange={e=>setD(p=>({...p,name:e.target.value}))} placeholder="Enter name or ID"/></F>
            <F lbl="Unit"><Seg opts={["Hour","KM"]} val={d.unit||"Hour"} onChange={v=>setD(p=>({...p,unit:v}))} small/></F>
            <F lbl={`${d.unit||"Hour"}s Used`}><Inp type="number" value={d.hours} onChange={e=>setD(p=>({...p,hours:e.target.value}))} placeholder="0.0" step="0.5"/></F>
            <F lbl="Fuel Issued (Litres)" col={2}><Inp type="number" value={d.fuel} onChange={e=>setD(p=>({...p,fuel:e.target.value}))} placeholder="0"/></F>
          </Grid>
          <div style={{display:"flex",gap:"8px",marginTop:"12px"}}>
            <button onClick={add} style={{flex:1,padding:"11px",borderRadius:"8px",border:"none",background:AM,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"700"}}>✓ Add Machine</button>
            <button onClick={()=>{setShow(false);setD(mkAsset());}} style={{padding:"11px 16px",borderRadius:"8px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"14px",color:"#6b7280"}}>Cancel</button>
          </div>
        </div>
      ):(
        <button onClick={()=>setShow(true)} style={{width:"100%",padding:"11px",borderRadius:"8px",border:"2px dashed "+AM,background:"#fffbeb",color:"#92400e",cursor:"pointer",fontSize:"14px",fontWeight:"600",marginTop:"6px",display:"flex",alignItems:"center",justifyContent:"center",gap:"8px"}}>
          <i className="ti ti-plus" style={{fontSize:"16px"}} aria-hidden/>Add Machine / Equipment
        </button>
      )}
    </div>
  );
}

// ─── LABOUR SECTION ───────────────────────────────────────────────────────────
function LabourSection({contractors,onChange,labourTypes}){
  const ltypes=labourTypes||LABOUR_TYPES;
  const [adding,setAdding]=useState(false);
  const [dname,setDname]=useState("");
  function addC(){if(!dname.trim())return;onChange([...contractors,{_k:uid(),name:dname.trim(),entries:[{_k:uid(),type:"Helper",nos:""}]}]);setDname("");setAdding(false);}
  function addEntry(ck){onChange(contractors.map(c=>c._k===ck?{...c,entries:[...c.entries,{_k:uid(),type:"Helper",nos:""}]}:c));}
  function updE(ck,ek,f,v){onChange(contractors.map(c=>c._k===ck?{...c,entries:c.entries.map(e=>e._k===ek?{...e,[f]:v}:e)}:c));}
  function remE(ck,ek){onChange(contractors.map(c=>c._k===ck?{...c,entries:c.entries.filter(e=>e._k!==ek)}:c).filter(c=>c.entries.length>0));}
  function remC(ck){onChange(contractors.filter(c=>c._k!==ck));}
  return(
    <div>
      {contractors.map(c=>(
        <div key={c._k} style={{border:"2px solid #86efac",borderRadius:"10px",marginBottom:"10px",overflow:"hidden"}}>
          <div style={{background:"#f0fdf4",padding:"10px 14px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <span style={{fontWeight:"700",fontSize:"14px",color:"#166534"}}>👤 {c.name}</span>
            <div style={{display:"flex",gap:"8px"}}>
              <button onClick={()=>addEntry(c._k)} style={{padding:"6px 12px",borderRadius:"6px",border:"1px solid "+GN,background:"#fff",color:GN,cursor:"pointer",fontSize:"12px",fontWeight:"600"}}>+ Labour Type</button>
              <button onClick={()=>remC(c._k)} style={{padding:"6px 10px",borderRadius:"6px",border:"1px solid #fca5a5",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"13px"}}>✕</button>
            </div>
          </div>
          <div style={{padding:"12px 14px",display:"flex",flexDirection:"column",gap:"8px"}}>
            {c.entries.map(e=>(
              <div key={e._k} style={{display:"grid",gridTemplateColumns:"1fr 90px 44px",gap:"8px",alignItems:"center"}}>
                <Sel val={e.type} onChange={v=>updE(c._k,e._k,"type",v)} opts={ltypes}/>
                <Inp type="number" value={e.nos} onChange={x=>updE(c._k,e._k,"nos",x.target.value)} placeholder="Nos" style={{textAlign:"center"}}/>
                <button onClick={()=>remE(c._k,e._k)} style={{width:"44px",height:"44px",borderRadius:"8px",border:"1px solid #fca5a5",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"16px",display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
              </div>
            ))}
          </div>
        </div>
      ))}
      {adding?(
        <div style={{background:"#f8fafc",border:"1px solid #d1d5db",borderRadius:"10px",padding:"14px",marginTop:"8px"}}>
          <L t="Contractor / Agency Name"/>
          <div style={{display:"flex",gap:"8px"}}>
            <Inp value={dname} onChange={e=>setDname(e.target.value)} placeholder="e.g. Anita Devi, Rajiv Kumar" onKeyDown={e=>e.key==="Enter"&&addC()} style={{flex:1}}/>
            <button onClick={addC} style={{padding:"11px 16px",borderRadius:"8px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"700",whiteSpace:"nowrap"}}>Add</button>
          </div>
          <button onClick={()=>setAdding(false)} style={{marginTop:"8px",padding:"8px 14px",borderRadius:"6px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"13px",color:"#6b7280"}}>Cancel</button>
        </div>
      ):(
        <button onClick={()=>setAdding(true)} style={{width:"100%",padding:"11px",borderRadius:"8px",border:"2px dashed "+GN,background:"#f0fdf4",color:"#166534",cursor:"pointer",fontSize:"14px",fontWeight:"600",marginTop:"6px",display:"flex",alignItems:"center",justifyContent:"center",gap:"8px"}}>
          + Add Contractor
        </button>
      )}
    </div>
  );
}

// ─── ENGINEER ROW ─────────────────────────────────────────────────────────────
function EngineerRow({e,onSave,onRemove,inchargeOpts,designationOpts,deptOpts}){
  const [ed,setEd]=useState(false);
  const [ic,setIc]=useState(e.incharge||"");
  const [des,setDes]=useState(e.designation||"");
  const [dept,setDept]=useState(e.dept||"");
  const [del,setDel]=useState(false);
  const mobile=useMobile();
  const icOpts=inchargeOpts||INCHARGE_OPTS;
  const desOpts=designationOpts||DESIGNATIONS;
  const dpOpts=deptOpts||DEPTS;
  // Sync if parent data changes (e.g. after Firebase update)
  useEffect(()=>{setIc(e.incharge||"");setDes(e.designation||"");setDept(e.dept||"");},[e.incharge,e.designation,e.dept]);
  return(
    <tr style={{borderBottom:"1px solid #f3f4f6"}}>
      {/* Name — read only, comes from Users tab */}
      <td style={{padding:"10px"}}>
        <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
          <Av name={e.name} sz={28}/>
          <span style={{fontWeight:"700"}}>{e.name}</span>
        </div>
      </td>
      {/* Dept — always a select in edit, badge in read */}
      {!mobile&&<td style={{padding:"10px"}}>
        {ed
          ?<select value={dept} onChange={e=>setDept(e.target.value)} style={{padding:"7px 8px",borderRadius:"6px",border:"1px solid #d1d5db",fontSize:"12px",background:"#fff",width:"100%"}}>
              <option value="">— Select —</option>
              {dpOpts.map(o=><option key={o}>{o}</option>)}
            </select>
          :<DeptB dept={e.dept}/>}
      </td>}
      {/* Designation */}
      {!mobile&&<td style={{padding:"10px",color:"#6b7280",fontSize:"12px"}}>
        {ed
          ?<select value={des} onChange={e=>setDes(e.target.value)} style={{padding:"7px 8px",borderRadius:"6px",border:"1px solid #d1d5db",fontSize:"12px",background:"#fff",width:"100%"}}>
              <option value="">— Select —</option>
              {desOpts.map(o=><option key={o}>{o}</option>)}
            </select>
          :<span>{e.designation||"—"}</span>}
      </td>}
      {/* Incharge */}
      <td style={{padding:"10px"}}>
        {ed
          ?<select value={ic} onChange={e=>setIc(e.target.value)} style={{padding:"7px 8px",borderRadius:"6px",border:"1px solid #d1d5db",fontSize:"12px",background:"#fff",width:"100%"}}>
              <option value="">— Select —</option>
              {icOpts.map(o=><option key={o}>{o}</option>)}
            </select>
          :<span style={{color:"#6b7280",fontSize:"12px"}}>{e.incharge||<span style={{color:"#f59e0b"}}>Not set</span>}</span>}
      </td>
      {/* Actions */}
      <td style={{padding:"10px"}}>
        {ed?(
          <div style={{display:"flex",gap:"6px"}}>
            <button onClick={()=>{onSave(e.id,{incharge:ic,designation:des,dept});setEd(false);}} style={{padding:"8px 14px",borderRadius:"6px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700"}}>Save</button>
            <button onClick={()=>{setEd(false);setIc(e.incharge||"");setDes(e.designation||"");setDept(e.dept||"");}} style={{padding:"8px 12px",borderRadius:"6px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"13px"}}>✕</button>
          </div>
        ):(
          <div style={{display:"flex",gap:"6px",flexWrap:"wrap"}}>
            <button onClick={()=>setEd(true)} style={{padding:"7px 12px",borderRadius:"6px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"12px",color:NV,fontWeight:"600"}}>Edit</button>
            {del
              ?<><button onClick={()=>onRemove(e.id)} style={{padding:"7px 10px",borderRadius:"6px",border:"none",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"12px",fontWeight:"700"}}>Confirm</button>
                  <button onClick={()=>setDel(false)} style={{padding:"7px 10px",borderRadius:"6px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"12px"}}>No</button></>
              :<button onClick={()=>setDel(true)} style={{padding:"7px 10px",borderRadius:"6px",border:"1px solid #fca5a5",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"12px"}}>Remove</button>
            }
          </div>
        )}
      </td>
    </tr>
  );
}

// ─── MATERIAL TX FORM ─────────────────────────────────────────────────────────
function MatTxNew({onAdd,editData,onEditSave,onEditCancel,lists,onDirty,requiredFields}){
  const [tx,setTx]=useState(editData||mkMatTx());
  const [showMach,setShowMach]=useState((editData?.assets||[]).length>0);
  const [showLabour,setShowLabour]=useState((editData?.contractors||[]).length>0);
  const mobile=useMobile();
  const isEdit=!!editData;
  const rf=requiredFields||{};
  const matUnits=lists?.units||["Cum","MT","Nos","Kg","Rmt"];
  const bulkMats=lists?.bulkMats||BULK_MATS;
  function setField(f,v){
    if(onDirty&&!isEdit)onDirty();
    setTx(p=>({...p,[f]:v}));
  }
  const u=(f)=>(v)=>setField(f,v);
  function handleAdd(){
    if(!tx.material){alert("Please select a Material first.");return;}
    const isRecv=tx.txType==="receive";
    const missing=[];
    if(isRecv){
      if(rf.recvCH&&!tx.recvCH)missing.push("Receiving Chainage");
      if(rf.source&&!tx.source)missing.push("Source / Origin");
      if(rf.recvQty&&!tx.recvQty)missing.push("Received Quantity");
      if(rf.recvTransporter&&!tx.recvTransporter)missing.push("Transporter");
      if(rf.recvLoads&&!tx.recvLoads)missing.push("No. of Loads");
    } else {
      if(rf.sendFromCH&&!tx.sendFromCH)missing.push("Sending From Chainage");
      if(rf.sendToCH&&!tx.sendToCH)missing.push("Sending To Chainage");
      if(rf.sendQty&&!tx.sendQty)missing.push("Sent Quantity");
      if(rf.sendTransporter&&!tx.sendTransporter)missing.push("Transporter");
      if(rf.sendLoads&&!tx.sendLoads)missing.push("No. of Loads");
    }
    if(rf.matRemarks&&!tx.remarks)missing.push("Remarks");
    if(missing.length>0){alert("Please fill required fields:\n• "+missing.join("\n• "));return;}
    if(isEdit){onEditSave({...tx});}else{onAdd({...tx,_k:tx._k||uid()});setTx(mkMatTx());}
  }
  return(
    <div style={{background:"#fff",border:`2px ${isEdit?"solid #0f766e":"dashed #14b8a6"}`,borderRadius:"12px",padding:"16px",marginTop:"8px"}}>
      <div style={{fontWeight:"700",fontSize:"15px",color:"#0f766e",marginBottom:"14px"}}>{isEdit?"✏️ Edit Material Entry":"📦 New Material Entry"}</div>
      <Grid cols={mobile?"1fr":"1fr 1fr"}>
        <F lbl="Transaction Type" req><Seg opts={["receive","send"]} val={tx.txType} onChange={u("txType")}/></F>
        <F lbl="Material" req>
          <select value={tx.material} onChange={e=>{setField('material',e.target.value);setTx(p=>({...p,matCustom:''}))}} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px",color:"#111827",background:"#fff"}}>
            <option value="">— Select Material —</option>
            <optgroup label="Materials">{bulkMats.filter(m=>m!=="Other").map(m=><option key={m}>{m}</option>)}</optgroup>
            <option value="Other">Other (type below)</option>
          </select>
          {tx.material==="Other"&&<Inp value={tx.matCustom} onChange={e=>setField('matCustom',e.target.value)} placeholder="Enter material name" style={{marginTop:"8px"}}/>}
        </F>
      </Grid>
      {tx.txType==="receive"&&(
        <div style={{background:"#f0fdf4",border:"2px solid #86efac",borderRadius:"10px",padding:"14px",marginBottom:"14px",marginTop:"6px"}}>
          <div style={{fontWeight:"700",fontSize:"13px",color:GN,marginBottom:"12px"}}>📥 Receiving Details</div>
          <Grid cols={mobile?"1fr":"1fr 1fr 1fr"}>
            <F lbl="Receiving Chainage" req={rf.recvCH}><Inp value={tx.recvCH} onChange={e=>setField('recvCH',e.target.value)} placeholder="e.g. 39+335"/></F>
            <F lbl="Source / Origin" req={rf.source}><Inp value={tx.source} onChange={e=>setField('source',e.target.value)} placeholder="Quarry name or CH"/></F>
            <F lbl="Quantity"><Inp type="number" value={tx.recvQty} onChange={e=>setField('recvQty',e.target.value)} placeholder="0.000"/></F>
            <F lbl="Unit"><Sel val={tx.recvUnit} onChange={u("recvUnit")} opts={matUnits}/></F>
            <F lbl="Transporter"><Inp value={tx.recvTransporter} onChange={e=>setField('recvTransporter',e.target.value)} placeholder="Name"/></F>
            <F lbl="No. of Loads"><Inp type="number" value={tx.recvLoads} onChange={e=>setField('recvLoads',e.target.value)} placeholder="0"/></F>
          </Grid>
        </div>
      )}
      {tx.txType==="send"&&(
        <div style={{background:"#eff6ff",border:"2px solid #93c5fd",borderRadius:"10px",padding:"14px",marginBottom:"14px",marginTop:"6px"}}>
          <div style={{fontWeight:"700",fontSize:"13px",color:"#1d4ed8",marginBottom:"12px"}}>📤 Sending Details</div>
          <Grid cols={mobile?"1fr":"1fr 1fr 1fr"}>
            <F lbl="Sending From Chainage" req={rf.sendFromCH}><Inp value={tx.sendFromCH} onChange={e=>setField('sendFromCH',e.target.value)} placeholder="e.g. 35+000"/></F>
            <F lbl="Sending To Chainage / Location" req={rf.sendToCH}><Inp value={tx.sendToCH} onChange={e=>setField('sendToCH',e.target.value)} placeholder="e.g. 39+335 or Stock Yard"/></F>
            <F lbl="Quantity"><Inp type="number" value={tx.sendQty} onChange={e=>setField('sendQty',e.target.value)} placeholder="0.000"/></F>
            <F lbl="Unit"><Sel val={tx.sendUnit} onChange={u("sendUnit")} opts={matUnits}/></F>
            <F lbl="Transporter"><Inp value={tx.sendTransporter} onChange={e=>setField('sendTransporter',e.target.value)} placeholder="Name"/></F>
            <F lbl="No. of Loads"><Inp type="number" value={tx.sendLoads} onChange={e=>setField('sendLoads',e.target.value)} placeholder="0"/></F>
          </Grid>
        </div>
      )}
      <F lbl="Remarks"><Inp value={tx.remarks} onChange={e=>setField('remarks',e.target.value)} placeholder="Any notes"/></F>
      {/* Accordion-style buttons for optional sections */}
      <button onClick={()=>setShowMach(s=>!s)} style={{width:"100%",marginTop:"10px",padding:"12px 16px",borderRadius:"10px",border:`2px solid ${showMach?"#b45309":"#d1d5db"}`,background:showMach?"#fffbeb":"#f9fafb",cursor:"pointer",fontSize:"14px",fontWeight:"700",color:showMach?"#92400e":"#374151",display:"flex",alignItems:"center",justifyContent:"space-between",gap:"8px"}}>
        <span>⚙️ Machines used for this movement {(tx.assets||[]).length>0&&<span style={{background:"#d97706",color:"#fff",borderRadius:"10px",padding:"1px 8px",fontSize:"11px",marginLeft:"6px"}}>{tx.assets.length}</span>}</span>
        <i className={`ti ti-chevron-${showMach?"up":"down"}`} style={{fontSize:"16px"}}/>
      </button>
      {showMach&&<div style={{padding:"12px 0 0"}}><AssetSection assets={tx.assets} onChange={v=>setTx(p=>({...p,assets:v}))} assetGroups={lists?.assetGroups}/></div>}
      <button onClick={()=>setShowLabour(s=>!s)} style={{width:"100%",marginTop:"8px",padding:"12px 16px",borderRadius:"10px",border:`2px solid ${showLabour?"#166534":"#d1d5db"}`,background:showLabour?"#f0fdf4":"#f9fafb",cursor:"pointer",fontSize:"14px",fontWeight:"700",color:showLabour?"#166534":"#374151",display:"flex",alignItems:"center",justifyContent:"space-between",gap:"8px"}}>
        <span>👷 Labour used for this movement {(tx.contractors||[]).length>0&&<span style={{background:GN,color:"#fff",borderRadius:"10px",padding:"1px 8px",fontSize:"11px",marginLeft:"6px"}}>{tx.contractors.length} contractor(s)</span>}</span>
        <i className={`ti ti-chevron-${showLabour?"up":"down"}`} style={{fontSize:"16px"}}/>
      </button>
      {showLabour&&<div style={{padding:"12px 0 0"}}><LabourSection contractors={tx.contractors} onChange={v=>setTx(p=>({...p,contractors:v}))} labourTypes={lists?.labourTypes}/></div>}
      <div style={{display:"flex",gap:"8px",marginTop:"14px"}}>
        <button onClick={handleAdd} style={{flex:1,padding:"14px",borderRadius:"10px",border:"none",background:"#0f766e",color:"#fff",cursor:"pointer",fontSize:"15px",fontWeight:"800",display:"flex",alignItems:"center",justifyContent:"center",gap:"8px"}}>
          {isEdit?"💾 Save Changes":"📦 Add This Material Entry"}
        </button>
        {isEdit&&<button onClick={onEditCancel} style={{padding:"14px 18px",borderRadius:"10px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"14px",color:"#6b7280"}}>Cancel</button>}
      </div>
    </div>
  );
}

// ─── ACTIVITY NEW ─────────────────────────────────────────────────────────────
function ActivityNew({onAdd,wt,initialData,editMode,lists,onDirty,requiredFields}){
  const [a,setA]=useState(initialData||mkAct());
  const mobile=useMobile();
  const rf=requiredFields||{};
  function upd(f,v){
    if(onDirty&&!editMode)onDirty();
    setA(prev=>{const n={...prev,[f]:v};if(["nos","length","width","depth","unit"].includes(f)){const t=calcTheo({...n});if(t)n.theoQty=t;}return n;});
  }
  function handleAdd(){
    if(!a.actType&&!a.actCustom){alert("Please select a Work Type first.");return;}
    const missing=[];
    if(rf.actDesc&&!a.desc)missing.push("Description");
    if(rf.chFrom&&!a.chFrom)missing.push("Chainage From");
    if(rf.chTo&&!a.chTo)missing.push("Chainage To");
    if(rf.side&&!a.side)missing.push("Side");
    if(rf.theoQty&&!a.theoQty&&!a.prodQty)missing.push("Quantity");
    if(rf.prodQty&&a.prodMat&&!a.prodQty)missing.push("Production Quantity");
    if(rf.actRemarks&&!a.remarks)missing.push("Remarks");
    if(missing.length>0){alert("Please fill required fields:\n• "+missing.join("\n• "));return;}
    onAdd({...a,_k:a._k||uid()});
    if(!editMode)setA(mkAct());
  }
  const sideOpts=lists?.sides||SIDES;
  const unitOpts=lists?.units||UNITS;
  // Parse prodMats from lists (format: "Name|Unit") or use constant
  // prodMats: supports both "Name" (new) and "Name|Unit" (old) formats
  const parsedProdMats=(lists?.prodMats||[]).map(s=>{if(!s)return null;const parts=(s||"").split("|");const n=parts[0].trim();const u=parts[1]?parts[1].trim():"Cum";return n?{n,u}:null;}).filter(Boolean);
  const prodMatsOpts=parsedProdMats.length>0?parsedProdMats:PROD_MATS;
  function calcTheo(x){const n=parseFloat(x.nos)||1,l=parseFloat(x.length)||0,w=parseFloat(x.width)||0,d=parseFloat(x.depth)||0;if(x.unit==="Cum"&&l&&w&&d)return(n*l*w*d).toFixed(3);if(x.unit==="Sqm"&&l&&w)return(n*l*w).toFixed(3);if(x.unit==="Rm"&&l)return(n*l).toFixed(3);if(x.unit==="Nos")return String(n);return "";}
  const pmObj=prodMatsOpts.find(x=>x.n===a.prodMat);
  return(
    <div>
      <SB emoji="🔧" title="1 — Work Type  (Select This First!)" sub="What type of work did you do?" color={PU}>
        <Grid cols="1fr">
          <F lbl="Work Type" req><WTInput value={a.actType} onChange={v=>setA(p=>({...p,actType:v}))} wt={wt}/>{a.actType==="Other"&&<Inp value={a.actCustom} onChange={e=>setA(p=>({...p,actCustom:e.target.value}))} placeholder="Specify work type" style={{marginTop:"8px"}}/>}</F>
          <F lbl="Description of Work" req={rf.actDesc}><Inp value={a.desc} onChange={e=>upd("desc",e.target.value)} placeholder="e.g. Embankment 3rd Layer Completed"/></F>
        </Grid>
      </SB>
      <SB emoji="📍" title="2 — Location (Chainage & Side)" color="#0f766e">
        <Grid cols={mobile?"1fr 1fr":"1fr 1fr 1fr 1fr"}>
          <F lbl="From Chainage" req={rf.chFrom}><Inp value={a.chFrom} onChange={e=>upd("chFrom",e.target.value)} placeholder="e.g. 39+335"/></F>
          <F lbl="To Chainage" req={rf.chTo}><Inp value={a.chTo} onChange={e=>upd("chTo",e.target.value)} placeholder="e.g. 39+340"/></F>
          <F lbl="Side" req={rf.side}><Sel val={a.side} onChange={v=>upd("side",v)} opts={sideOpts}/></F>
          <F lbl="CW / Area"><Inp value={a.cw} onChange={e=>upd("cw",e.target.value)} placeholder="MCW / SR / RE…"/></F>
        </Grid>
      </SB>
      <SB emoji="📐" title="3 — Measurement & Quantity" color="#0369a1">
        <Grid cols={mobile?"1fr 1fr":"1fr 1fr 1fr 1fr 1fr"}>
          <F lbl="Unit"><Sel val={a.unit} onChange={v=>upd("unit",v)} opts={unitOpts}/></F>
          <F lbl="Nos"><Inp type="number" value={a.nos} onChange={e=>upd("nos",e.target.value)} min="1"/></F>
          <F lbl="Length (m)"><Inp type="number" value={a.length} onChange={e=>upd("length",e.target.value)} placeholder="0.000" step="0.001"/></F>
          <F lbl="Width (m)"><Inp type="number" value={a.width} onChange={e=>upd("width",e.target.value)} placeholder="0.000" step="0.001"/></F>
          <F lbl="Depth (m)"><Inp type="number" value={a.depth} onChange={e=>upd("depth",e.target.value)} placeholder="0.000" step="0.001"/></F>
        </Grid>
        <div style={{marginTop:"14px"}}><Grid cols={mobile?"1fr":"1fr 1fr"}>
          <F lbl="Theoretical Qty" req={rf.theoQty}><Inp value={a.theoQty} onChange={e=>upd("theoQty",e.target.value)} placeholder="Auto-calculated" style={{background:"#f8fafc"}}/></F>
          <F lbl="Remarks"><Inp value={a.remarks} onChange={e=>upd("remarks",e.target.value)} placeholder="Any notes"/></F>
        </Grid></div>
        <div style={{background:"#fffbeb",border:"2px solid #fde68a",borderRadius:"10px",padding:"14px",marginTop:"14px"}}>
          <div style={{fontWeight:"700",fontSize:"13px",color:"#92400e",marginBottom:"12px"}}>🏭 Production Material Used <span style={{fontWeight:"400",fontSize:"11px"}}>— only if RMC / DBM / BC / WMM was consumed</span></div>
          <Grid cols={mobile?"1fr":"2fr 1fr 1fr"}>
            <F lbl="Material Used">
              <select value={a.prodMat} onChange={e=>{const m=prodMatsOpts.find(x=>x.n===e.target.value);setA(p=>({...p,prodMat:e.target.value,prodMatCustom:"",unit:m?m.u:p.unit,prodQty:""}));}} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px",background:"#fff",color:"#111827"}}>
                <option value="">— Not applicable —</option>
                {prodMatsOpts.map(x=><option key={x.n}>{x.n}</option>)}
              </select>
            </F>
            <F lbl={pmObj?"Qty Used":"Qty Used"}><Inp type="number" value={a.prodQty} onChange={e=>upd("prodQty",e.target.value)} placeholder="0.000" step="0.001" disabled={!a.prodMat}/></F>
            <F lbl="Unit"><Sel val={a.prodUnit||pmObj?.u||"Cum"} onChange={v=>setA(p=>({...p,prodUnit:v}))} opts={unitOpts}/></F>
          </Grid>
        </div>
      </SB>
      <SB emoji="⚙️" title="4 — Plant & Machinery Used" sub="Add each machine used for THIS activity" color="#b45309">
        <AssetSection assets={a.assets} onChange={v=>setA(p=>({...p,assets:v}))} assetGroups={lists?.assetGroups}/>
      </SB>
      <SB emoji="👷" title="5 — Labour & Contractors" sub="Add contractor once — then add multiple labour types under them" color={GN}>
        <LabourSection contractors={a.contractors} onChange={v=>setA(p=>({...p,contractors:v}))} labourTypes={lists?.labourTypes}/>
      </SB>
      <button onClick={handleAdd} style={{width:"100%",padding:"16px",borderRadius:"12px",border:`3px ${editMode?"solid":"dashed"} ${AM}`,background:"#fffbeb",color:"#92400e",cursor:"pointer",fontSize:"16px",fontWeight:"800",display:"flex",alignItems:"center",justifyContent:"center",gap:"10px",marginTop:"4px"}}>
        <i className="ti ti-circle-plus" style={{fontSize:"22px"}} aria-hidden/>{editMode?"💾 Save Changes":"Add This Activity to My List"}
      </button>
    </div>
  );
}

// ─── EXCEL ────────────────────────────────────────────────────────────────────
function doPrintDPR(dayS,engineers,date,projectName){
  const rows=engineers.map(e=>{
    const ss=dayS.filter(s=>s.engineer===e.name);
    if(ss.length===0)return{...e,absent:true,subs:[]};
    return{...e,absent:false,subs:ss};
  });
  const acts=dayS.flatMap(s=>(s.activities||[]).map(a=>({...a,_eng:s.engineer,_shift:s.shift,_ok:s.approved})));
  const mats=dayS.flatMap(s=>(s.matTxs||[]).map(m=>({...m,_eng:s.engineer})));
  const dt=new Date(date+"T12:00:00").toLocaleDateString("en-IN",{weekday:"long",day:"2-digit",month:"long",year:"numeric"});
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>DPR — ${date}</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0;}
    body{font-family:Arial,sans-serif;padding:24px 32px;color:#111;font-size:12px;}
    h1{font-size:18px;color:#1e3a5f;margin-bottom:2px;}
    .sub{font-size:11px;color:#666;margin-bottom:18px;}
    table{width:100%;border-collapse:collapse;margin-bottom:18px;font-size:11px;}
    th{background:#1e3a5f;color:#fff;padding:7px 10px;text-align:left;}
    td{padding:7px 10px;border-bottom:1px solid #e5e7eb;vertical-align:top;}
    tr:nth-child(even) td{background:#f8fafc;}
    .sec{font-size:13px;font-weight:700;color:#1e3a5f;margin:16px 0 8px;border-bottom:2px solid #1e3a5f20;padding-bottom:4px;}
    .badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;}
    .ok{background:#dcfce7;color:#166534;}
    .pend{background:#fef3c7;color:#92400e;}
    .absent{color:#dc2626;}
    @media print{body{padding:12px 18px;}button{display:none;}}
  </style></head><body>
  <h1>Daily Progress Report — ${projectName}</h1>
  <div class="sub">${dt} · Generated ${new Date().toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"})}</div>
  
  <div class="sec">Attendance Summary</div>
  <table>
    <thead><tr><th>#</th><th>Engineer</th><th>Dept</th><th>Incharge</th><th>Shift</th><th>Activities</th><th>Mat.Moves</th><th>Status</th></tr></thead>
    <tbody>${rows.map((e,i)=>e.absent
      ?`<tr><td>${i+1}</td><td>${e.name}</td><td>${e.dept||"—"}</td><td>${e.incharge||"—"}</td><td>—</td><td>—</td><td>—</td><td class="absent">Absent</td></tr>`
      :e.subs.map(s=>`<tr><td>${i+1}</td><td><strong>${e.name}</strong></td><td>${e.dept||"—"}</td><td>${e.incharge||"—"}</td><td>${s.shift}</td><td>${(s.activities||[]).length}</td><td>${(s.matTxs||[]).length}</td><td><span class="${s.approved?"ok":"pend"}">${s.approved?"✓ Approved by "+s.approvedBy:"⏳ Pending"}</span></td></tr>`).join("")
    ).join("")}</tbody>
  </table>

  ${acts.length>0?`<div class="sec">Work Activities (${acts.length})</div>
  <table>
    <thead><tr><th>Engineer</th><th>Work Type</th><th>Description</th><th>From CH</th><th>To CH</th><th>Side</th><th>Qty</th><th>Unit</th></tr></thead>
    <tbody>${acts.map(a=>`<tr><td>${a._eng}</td><td>${a.actType||(a.actCustom||"—")}</td><td>${a.desc||"—"}</td><td>${a.chFrom||"—"}</td><td>${a.chTo||"—"}</td><td>${a.side||"—"}</td><td>${a.theoQty||"—"}</td><td>${a.unit||"—"}</td></tr>`).join("")}</tbody>
  </table>`:""}

  ${mats.length>0?`<div class="sec">Material Movements (${mats.length})</div>
  <table>
    <thead><tr><th>Engineer</th><th>Material</th><th>Type</th><th>CH</th><th>Qty</th><th>Unit</th><th>Transporter</th></tr></thead>
    <tbody>${mats.map(m=>`<tr><td>${m._eng}</td><td>${m.material==="Other"?m.matCustom:m.material}</td><td>${m.txType}</td><td>${m.recvCH||m.sendFromCH||"—"}</td><td>${m.recvQty||m.sendQty||"—"}</td><td>${m.recvUnit||m.sendUnit||"—"}</td><td>${m.recvTransporter||m.sendTransporter||"—"}</td></tr>`).join("")}</tbody>
  </table>`:""}

  <div style="margin-top:40px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;font-size:11px;">
    <div><div style="border-top:1px solid #666;padding-top:6px;text-align:center;">Prepared By</div></div>
    <div><div style="border-top:1px solid #666;padding-top:6px;text-align:center;">Checked By</div></div>
    <div><div style="border-top:1px solid #666;padding-top:6px;text-align:center;">Approved By</div></div>
  </div>
  <script>window.onload=()=>window.print();<\/script>
  </body></html>`;
  const w=window.open("","_blank");
  w.document.write(html);
  w.document.close();
}

function doExcel(subs,engineers,date,projectName){
  // date can be "YYYY-MM-DD" (single day) or "YYYY-MM-DD to YYYY-MM-DD" (range from reports tab)
  let F;
  if(date&&date.includes(' to ')){
    const [from,to]=date.split(' to ');
    F=subs.filter(s=>s.date>=from&&s.date<=to);
  } else {
    F=subs.filter(s=>s.date===date);
  }
  const aR=[["Date","Engineer","Dept","Shift","Incharge","Work Type","Desc","From CH","To CH","Side","CW","Unit","Nos","Length","Width","Depth","Theo Qty","Prod Mat","Prod Qty","Diff","Remarks","Approved","Approved By"]];
  F.forEach(s=>(s.activities||[]).forEach(a=>aR.push([s.date,s.engineer,s.dept,s.shift,s.incharge,a.actType||(a.actCustom||""),a.desc,a.chFrom,a.chTo,a.side,a.cw,a.unit,a.nos,a.length,a.width,a.depth,a.theoQty,a.prodMat,a.prodQty,a.remarks,s.approved?"Yes":"No",s.approvedBy||"—"])));
  const mR=[["Date","Engineer","Type","Material","Recv CH","Source","Recv Qty","Unit","Transporter","Loads","Send From","Send To","Send Qty","Send Unit","Send Transporter","Send Loads","Remarks"]];
  F.forEach(s=>(s.matTxs||[]).forEach(t=>{const mt=t.material==="Other"?t.matCustom:t.material;mR.push([s.date,s.engineer,t.txType,mt,t.recvCH,t.source,t.recvQty,t.recvUnit,t.recvTransporter,t.recvLoads,t.sendFromCH,t.sendToCH,t.sendQty,t.sendUnit,t.sendTransporter,t.sendLoads,t.remarks]);}));
  const pR=[["Date","Engineer","Work Type","Asset Group","Asset Name","Hours","Fuel (L)"]];
  F.forEach(s=>(s.activities||[]).forEach(a=>(a.assets||[]).forEach(x=>pR.push([s.date,s.engineer,a.actType,x.group,x.name,x.hours,x.fuel]))));
  const lR=[["Date","Engineer","Work Type","Contractor","Labour Type","Nos"]];
  F.forEach(s=>(s.activities||[]).forEach(a=>(a.contractors||[]).forEach(c=>(c.entries||[]).forEach(e=>lR.push([s.date,s.engineer,a.actType,c.name,e.type,e.nos])))));
  const attR=[["Engineer","Dept","Incharge","Status","Shift","Activities","Approved"]];
  engineers.forEach(e=>{const ss=F.filter(s=>s.engineer===e.name);const pres=ss.length>0;attR.push([e.name,e.dept,e.incharge,pres?"Present":"Absent",[...new Set(ss.map(s=>s.shift))].join("/"),(ss.reduce((a,s)=>a+(s.activities||[]).length,0)),ss.filter(s=>s.approved).length+"/"+ss.length]);});
  const wb=XLSX.utils.book_new();
  [[aR,"Work Activities"],[mR,"Material Transactions"],[pR,"Plant & Machinery"],[lR,"Labour"],[attR,"Attendance"]].forEach(([rows,name])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),name));
  const isRange=date&&date.includes(' to ');
  const fileName=isRange
    ?`Consolidated_DPR_${date.replace(' to ','_to_')}.xlsx`
    :`Daily_DPR_${date}.xlsx`;
  XLSX.writeFile(wb,fileName);
}

// ─── LOGIN MODAL ──────────────────────────────────────────────────────────────
function LoginModal({users,onLogin,onClose}){
  const [pin,setPin]=useState("");
  const [err,setErr]=useState("");
  const [show,setShow]=useState(false);
  function tryLogin(){
    const p=pin.trim();
    // Hardcoded super-admin backup PIN — always works, even on an empty database
    if(p==="005566"){onLogin({id:"superadmin",name:"Super Admin",role:"admin",pin:"005566",caps:ROLE_CAPS.admin});return;}
    const u=users.find(x=>x.pin===p);if(u)onLogin(u);else{setErr("Wrong PIN. Please try again.");setPin("");}
  }
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.65)",zIndex:10000,display:"flex",alignItems:"flex-end",justifyContent:"center",padding:"0"}} onClick={onClose}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:"20px 20px 0 0",padding:"28px 24px 36px",maxWidth:"480px",width:"100%",boxShadow:"0 -8px 40px rgba(0,0,0,.25)"}}>
        <div style={{width:"40px",height:"4px",background:"#e5e7eb",borderRadius:"2px",margin:"0 auto 20px"}}/>
        <div style={{textAlign:"center",marginBottom:"24px"}}><div style={{fontSize:"32px",marginBottom:"10px"}}>🔐</div><div style={{fontSize:"18px",fontWeight:"800",color:NV}}>Sign In</div><div style={{fontSize:"13px",color:"#6b7280",marginTop:"4px"}}>Enter your PIN to access your role</div></div>
        <L t="Your PIN"/>
        <div style={{position:"relative",marginBottom:"18px"}}>
          <Inp type={show?"text":"password"} value={pin} onChange={e=>setPin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&tryLogin()} placeholder="Enter PIN" maxLength={8} style={{letterSpacing:show?"normal":"0.4em",fontSize:"20px",paddingRight:"48px",border:`2px solid ${err?"#ef4444":"#d1d5db"}`,textAlign:"center"}}/>
          <button onClick={()=>setShow(s=>!s)} style={{position:"absolute",right:"14px",top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:"#6b7280",fontSize:"18px"}}><i className={`ti ti-eye${show?"-off":""}`} aria-hidden/></button>
        </div>
        {err&&<div style={{color:RD,fontSize:"13px",marginBottom:"14px",fontWeight:"600",textAlign:"center"}}>⚠ {err}</div>}
        <button onClick={tryLogin} style={{width:"100%",padding:"15px",borderRadius:"12px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"16px",fontWeight:"800"}}>Sign In →</button>
      </div>
    </div>
  );
}

// ─── GLOBAL ENGINEERS PANEL ──────────────────────────────────────────────────
// ─── GLOBAL SETTINGS PANEL ───────────────────────────────────────────────────
function GlobalSettingsPanel({globalLists,setGlobalLists,flash,mobile}){
  const [edits,setEdits]=useState({});
  const rolesVal=edits.roles!==undefined?edits.roles:(globalLists.roles||Object.keys(ROLE_CAPS)).join("\n");
  const lockVal=edits.dateLockDays!==undefined?edits.dateLockDays:String(globalLists.dateLockDays??2);
  const starStartVal=edits.starStart!==undefined?edits.starStart:String(globalLists.starStart??5);
  const lateEngVal=edits.lateEngDeduct!==undefined?edits.lateEngDeduct:String(globalLists.lateEngDeduct??0.5);
  const lateIcVal=edits.lateInchargeDeduct!==undefined?edits.lateInchargeDeduct:String(globalLists.lateInchargeDeduct??0.25);
  const apvVal=edits.dprApprovalDays!==undefined?edits.dprApprovalDays:String(globalLists.dprApprovalDays??3);
  const rf=globalLists.requiredFields||{};

  // Field definitions grouped by section
  const FIELD_GROUPS=[
    {label:"📋 Step 1 — DPR Header",fields:[
      {k:"dept",label:"Department",note:"Auto-filled for engineers"},
      {k:"incharge",label:"Incharge / HOD",note:"Auto-filled for engineers"},
      {k:"weather",label:"Weather Condition"},
      {k:"difficulty",label:"Issues / Problems Faced"},
    ]},
    {label:"🔧 Step 3 — Work Activities",fields:[
      {k:"theoQty",label:"Theoretical Quantity",note:"Recommended ON"},
      {k:"chFrom",label:"Chainage From"},
      {k:"chTo",label:"Chainage To"},
      {k:"side",label:"Side (RHS/LHS/Both)"},
      {k:"actDesc",label:"Activity Description"},
      {k:"prodQty",label:"Production Quantity"},
      {k:"actRemarks",label:"Activity Remarks"},
    ]},
    {label:"📦 Step 2 — Material Movements (Receive)",fields:[
      {k:"recvQty",label:"Received Quantity",note:"Recommended ON"},
      {k:"recvCH",label:"Receiving Chainage"},
      {k:"source",label:"Source / Origin"},
      {k:"recvTransporter",label:"Transporter Name"},
      {k:"recvLoads",label:"No. of Loads"},
    ]},
    {label:"📦 Step 2 — Material Movements (Send)",fields:[
      {k:"sendQty",label:"Sent Quantity",note:"Recommended ON"},
      {k:"sendFromCH",label:"Sending From Chainage"},
      {k:"sendToCH",label:"Sending To Chainage"},
      {k:"sendTransporter",label:"Transporter Name"},
      {k:"sendLoads",label:"No. of Loads"},
    ]},
    {label:"📝 Remarks",fields:[
      {k:"matRemarks",label:"Material Remarks"},
      {k:"actRemarks",label:"Activity Remarks"},
    ]},
  ];

  function toggleField(k){
    const updated={...rf,[k]:!rf[k]};
    rtdbPut('config/globalLists',{...globalLists,requiredFields:updated})
      .then(()=>{setGlobalLists(p=>({...p,requiredFields:updated}));flash("✅ Saved");})
      .catch(e=>flash("Failed: "+e.message,"err"));
  }

  function saveRoles(){
    const lines=rolesVal.split("\n").map(s=>s.trim()).filter(Boolean);
    if(!lines.length){flash("List cannot be empty","err");return;}
    rtdbPut('config/globalLists',{...globalLists,roles:lines})
      .then(()=>{setGlobalLists(p=>({...p,roles:lines}));flash("✅ Roles saved");})
      .catch(e=>flash("Failed: "+e.message,"err"));
  }

  function saveWindowsAndStars(){
    const apv=parseInt(apvVal), lock=parseInt(lockVal);
    const starStart=parseFloat(starStartVal), engD=parseFloat(lateEngVal), icD=parseFloat(lateIcVal);
    if([apv,lock].some(n=>isNaN(n)||n<0)){flash("Enter 0 or a positive number of days","err");return;}
    if([starStart,engD,icD].some(n=>isNaN(n)||n<0)){flash("Enter valid non-negative star values","err");return;}
    rtdbPut('config/globalLists',{...globalLists,dprApprovalDays:apv,dateLockDays:lock,starStart,lateEngDeduct:engD,lateInchargeDeduct:icD})
      .then(()=>{setGlobalLists(p=>({...p,dprApprovalDays:apv,dateLockDays:lock,starStart,lateEngDeduct:engD,lateInchargeDeduct:icD}));flash("✅ Settings saved");})
      .catch(e=>flash("Failed: "+e.message,"err"));
  }

  return(
    <div style={{padding:mobile?"8px":"20px",maxWidth:"680px",margin:"0 auto",display:"flex",flexDirection:"column",gap:"14px"}}>

      {/* Required Fields */}
      <Card>
        <div style={{fontWeight:"700",fontSize:"15px",color:NV,marginBottom:"4px"}}>✅ Mandatory Fields in DPR Form</div>
        <p style={{fontSize:"13px",color:"#6b7280",marginBottom:"14px",marginTop:0}}>
          Toggle which fields engineers must fill before they can submit a DPR. Changes apply immediately to all users.
        </p>
        <div style={{display:"flex",flexDirection:"column",gap:"16px"}}>
          {FIELD_GROUPS.filter((g,i)=>i<4).map(g=>(
            <div key={g.label}>
              <div style={{fontWeight:"700",fontSize:"12px",color:"#374151",textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:"8px",paddingBottom:"6px",borderBottom:"1px solid #f3f4f6"}}>{g.label}</div>
              <div style={{display:"flex",flexDirection:"column",gap:"1px"}}>
                {g.fields.filter((f,i,arr)=>arr.findIndex(x=>x.k===f.k)===i).map(f=>{
                  const isOn=rf[f.k]??false;
                  return(
                    <label key={f.k} style={{display:"flex",alignItems:"center",gap:"12px",padding:"9px 10px",borderRadius:"8px",cursor:"pointer",background:isOn?"#eff6ff":"transparent",border:isOn?"1px solid #bfdbfe":"1px solid transparent",transition:"all .15s"}}>
                      <div onClick={()=>toggleField(f.k)} style={{width:"44px",height:"24px",borderRadius:"12px",background:isOn?"#3b82f6":"#d1d5db",position:"relative",flexShrink:0,cursor:"pointer",transition:"background .2s"}}>
                        <div style={{position:"absolute",top:"3px",left:isOn?"23px":"3px",width:"18px",height:"18px",borderRadius:"50%",background:"#fff",transition:"left .2s",boxShadow:"0 1px 3px rgba(0,0,0,.3)"}}/>
                      </div>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:"600",fontSize:"13px",color:isOn?"#1e40af":"#374151"}}>{f.label}</div>
                        {f.note&&<div style={{fontSize:"11px",color:"#9ca3af"}}>{f.note}</div>}
                      </div>
                      <span style={{fontSize:"11px",fontWeight:"700",color:isOn?"#1e40af":"#9ca3af",minWidth:"50px",textAlign:"right"}}>{isOn?"Required":"Optional"}</span>
                    </label>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
        <div style={{marginTop:"12px",background:"#fffbeb",borderRadius:"8px",padding:"10px 12px",fontSize:"12px",color:"#92400e"}}>
          ⚠️ Engineer name, Date, and Shift are always required and cannot be turned off.
        </div>
      </Card>

      {/* Approval windows and star rules */}
      <Card>
        <div style={{fontWeight:"700",fontSize:"15px",color:NV,marginBottom:"4px"}}>⚙️ Approval windows and star rules</div>
        <p style={{fontSize:"13px",color:"#6b7280",marginBottom:"16px",marginTop:0}}>Deadlines for approvals and back-dated DPRs, and how stars are deducted for late work.</p>

        <div style={{marginBottom:"16px"}}>
          <div style={{fontWeight:"700",fontSize:"13px",color:"#374151",marginBottom:"4px"}}>⏱ DPR Approval Window</div>
          <p style={{fontSize:"12px",color:"#6b7280",margin:"0 0 8px"}}>Time limit (in days) for an incharge to approve a DPR. Approving after this is logged as a backlog approval with star penalties.</p>
          <div style={{display:"flex",gap:"10px",alignItems:"center"}}>
            <input type="number" min="0" max="60" value={apvVal} onChange={e=>setEdits(p=>({...p,dprApprovalDays:e.target.value}))}
              style={{width:"100px",padding:"11px 12px",borderRadius:"8px",border:"2px solid #d1d5db",fontSize:"16px",fontWeight:"700",textAlign:"center"}}/>
            <span style={{fontSize:"13px",color:"#6b7280",fontWeight:"600"}}>days</span>
          </div>
        </div>

        <div style={{marginBottom:"16px"}}>
          <div style={{fontWeight:"700",fontSize:"13px",color:"#374151",marginBottom:"4px"}}>🔒 DPR Filling Lock</div>
          <p style={{fontSize:"12px",color:"#6b7280",margin:"0 0 8px"}}>Engineers can submit a DPR for up to this many days back. Beyond it, they must raise a late-entry request.</p>
          <div style={{display:"flex",gap:"10px",alignItems:"center"}}>
            <input type="number" min="0" max="30" value={lockVal} onChange={e=>setEdits(p=>({...p,dateLockDays:e.target.value}))}
              style={{width:"100px",padding:"11px 12px",borderRadius:"8px",border:"2px solid #d1d5db",fontSize:"16px",fontWeight:"700",textAlign:"center"}}/>
            <span style={{fontSize:"13px",color:"#6b7280",fontWeight:"600"}}>days (0 = no lock)</span>
          </div>
        </div>

        <div style={{marginBottom:"14px"}}>
          <div style={{fontWeight:"700",fontSize:"13px",color:"#374151",marginBottom:"4px"}}>⭐ Star Rating Rules</div>
          <p style={{fontSize:"12px",color:"#6b7280",margin:"0 0 8px"}}>Set the starting score and how many stars a late submission costs.</p>
          <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"repeat(3,1fr)",gap:"12px"}}>
            <div>
              <label style={{fontSize:"12px",fontWeight:"700",color:"#374151",display:"block",marginBottom:"6px"}}>Starting stars</label>
              <input type="number" min="0" max="10" step="0.5" value={starStartVal} onChange={e=>setEdits(p=>({...p,starStart:e.target.value}))} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"2px solid #d1d5db",fontSize:"15px",fontWeight:"700",textAlign:"center"}}/>
            </div>
            <div>
              <label style={{fontSize:"12px",fontWeight:"700",color:"#374151",display:"block",marginBottom:"6px"}}>Engineer penalty</label>
              <input type="number" min="0" max="5" step="0.05" value={lateEngVal} onChange={e=>setEdits(p=>({...p,lateEngDeduct:e.target.value}))} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"2px solid #d1d5db",fontSize:"15px",fontWeight:"700",textAlign:"center"}}/>
            </div>
            <div>
              <label style={{fontSize:"12px",fontWeight:"700",color:"#374151",display:"block",marginBottom:"6px"}}>Incharge penalty</label>
              <input type="number" min="0" max="5" step="0.05" value={lateIcVal} onChange={e=>setEdits(p=>({...p,lateInchargeDeduct:e.target.value}))} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"2px solid #d1d5db",fontSize:"15px",fontWeight:"700",textAlign:"center"}}/>
            </div>
          </div>
          <div style={{marginTop:"10px",fontSize:"12px",color:"#6b7280",background:"#f8fafc",borderRadius:"8px",padding:"9px 12px"}}>
            Example: an engineer who submits {Math.max(1,parseInt(lockVal)||0)} day{(Math.max(1,parseInt(lockVal)||0))===1?"":"s"} late drops from {parseFloat(starStartVal)||0}★ to {(Math.max(0,(parseFloat(starStartVal)||0)-(parseFloat(lateEngVal)||0))).toFixed(2)}★, and their incharge drops by {parseFloat(lateIcVal)||0}★.
          </div>
        </div>

        <button onClick={saveWindowsAndStars} style={{width:"100%",padding:"12px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"800"}}>💾 Save Settings</button>
      </Card>

      {/* Roles */}
      <Card>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"6px"}}>
          <div style={{fontWeight:"700",fontSize:"15px",color:NV}}>🎭 User Roles</div>
          <span style={{fontSize:"12px",color:"#9ca3af"}}>{(globalLists.roles||[]).length} roles</span>
        </div>
        <p style={{fontSize:"13px",color:"#6b7280",marginBottom:"10px",marginTop:0}}>One role per line. Appears in the Role dropdown when creating users.</p>
        <textarea value={rolesVal} onChange={e=>setEdits(p=>({...p,roles:e.target.value}))} rows={8}
          style={{width:"100%",boxSizing:"border-box",fontFamily:"monospace",fontSize:"14px",padding:"12px",border:"2px solid #d1d5db",borderRadius:"10px",background:"#f9fafb",lineHeight:"1.9",resize:"vertical"}}/>
        <div style={{display:"flex",gap:"10px",marginTop:"12px"}}>
          <button onClick={saveRoles} style={{flex:1,padding:"12px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"800"}}>💾 Save Roles</button>
          <button onClick={()=>setEdits(p=>({...p,roles:Object.keys(ROLE_CAPS).join("\n")}))} style={{padding:"12px 14px",borderRadius:"10px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"13px",color:"#6b7280"}}>Reset</button>
        </div>
      </Card>
    </div>
  );
}

// ─── USERS PANEL ─────────────────────────────────────────────────────────────
function UsersPanel({users,projects,mobile,newUsr,setNewUsr,addUser,reassignUser,saveUserCaps,flash,globalLists,user}){
  const [editId,setEditId]=useState(null);
  const [editData,setEditData]=useState({});
  const [userAudit,setUserAudit]=useState([]);
  const [showLog,setShowLog]=useState(false);
  useEffect(()=>{
    const r=ref(db,'userAudit');
    const cb=onValue(r,snap=>setUserAudit(toArr(snap.val()).sort((a,b)=>(b.ts||"").localeCompare(a.ts||""))));
    return()=>off(r,'value',cb);
  },[]);
  function logUserAudit(action,target,detail){
    const id=Date.now().toString(36)+Math.random().toString(36).slice(2,7);
    rtdbPut('userAudit/'+id,{id,ts:new Date().toISOString(),actor:user?.name||"Admin",action,target,detail:detail||""}).catch(()=>{});
  }
  const [search,setSearch]=useState("");
  const [roleFilter,setRoleFilter]=useState("all");
  const [projFilter,setProjFilter]=useState("all");
  const [groupBy,setGroupBy]=useState("none");
  const [delUser,setDelUser]=useState(null);
  const [delCode,setDelCode]=useState("");
  const [delInput,setDelInput]=useState("");
  const [showCreatePreview,setShowCreatePreview]=useState(false);
  const [previewTab,setPreviewTab]=useState("card");
  const [editDiff,setEditDiff]=useState(null);
  const roleOpts=globalLists?.roles||Object.keys(ROLE_CAPS);
  const PERMS=[["fill","✏️ Fill DPR"],["approve","✅ Approve"],["download","📥 Download"],["manage","👷 Manage Engineers"],["settings","⚙️ Edit Lists"]];
  const SH2={fontSize:"12px",fontWeight:"700",color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:"14px",paddingBottom:"10px",borderBottom:"1px solid #f3f4f6",display:"flex",alignItems:"center",justifyContent:"space-between"};

  function downloadUsersPDF(){
    const rows=users.map(u=>{
      const proj=projects.find(p=>p.id===u.assignedProjectId);
      const mgmtProjs=toArr(u.assignedProjectIds).map(pid=>projects.find(p=>p.id===pid)).filter(Boolean).map(p=>p.name).join(", ");
      const uc=u.caps||ROLE_CAPS[roleKey(u.role)]||ROLE_CAPS.engineer;
      const perms=PERMS.filter(([k])=>uc[k]).map(([,l])=>l.replace(/[^\w\s]/g,"").trim()).join(", ")||"—";
      return{name:u.name,role:u.role,pin:u.pin||"—",project:u.projectAccess==="all"?"All projects":(mgmtProjs||(proj?.name||"Not assigned")),perms};
    });
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>SPL Users List</title>
    <style>
      body{font-family:Arial,sans-serif;padding:32px;color:#111;}
      h1{font-size:20px;margin-bottom:4px;color:#1e3a5f;}
      .sub{font-size:13px;color:#6b7280;margin-bottom:24px;}
      table{width:100%;border-collapse:collapse;font-size:13px;}
      th{background:#1e3a5f;color:#fff;padding:10px 12px;text-align:left;font-size:12px;}
      td{padding:9px 12px;border-bottom:1px solid #e5e7eb;}
      tr:nth-child(even) td{background:#f8fafc;}
      .pin{font-family:monospace;font-weight:700;color:#dc2626;letter-spacing:0.1em;}
      @media print{body{padding:16px;}}
    </style></head><body>
    <h1>SPL Infrastructure — User Accounts</h1>
    <div class="sub">Generated ${new Date().toLocaleString("en-IN")} · Total ${users.length} users · <strong style="color:#dc2626">CONFIDENTIAL — Admin Use Only</strong></div>
    <table>
      <thead><tr><th>#</th><th>Name</th><th>Role</th><th>PIN</th><th>Project</th><th>Permissions</th></tr></thead>
      <tbody>${rows.map((r,i)=>`<tr><td>${i+1}</td><td><strong>${r.name}</strong></td><td>${r.role}</td><td class="pin">${r.pin}</td><td>${r.project}</td><td style="font-size:11px;color:#374151">${r.perms}</td></tr>`).join("")}</tbody>
    </table>
    <script>window.onload=()=>window.print();<\/script>
    </body></html>`;
    const w=window.open("","_blank");
    w.document.write(html);
    w.document.close();
  }

  function startEdit(u){setEditId(u.id);setEditData({name:u.name,role:u.role,pin:"",desc:u.desc||"",projectAccess:u.projectAccess||(u.assignedProjectId||toArr(u.assignedProjectIds).length?"specific":"none"),assignedProjectId:u.assignedProjectId||"",assignedProjectIds:u.assignedProjectIds||[],caps:u.caps||ROLE_CAPS[roleKey(u.role)]||ROLE_CAPS.engineer});}

  function saveEdit(u){
    const updated={...editData};
    if(!updated.pin)delete updated.pin;
    if(updated.projectAccess&&updated.projectAccess!=="specific"){updated.assignedProjectId="";updated.assignedProjectIds=[];}
    const pName=id=>projects.find(p=>p.id===id)?.name||"None";
    const changes=[];
    if(updated.name!==undefined&&updated.name!==u.name)changes.push({f:"Name",from:u.name,to:updated.name});
    if(updated.role!==undefined&&updated.role!==u.role)changes.push({f:"Role",from:u.role,to:updated.role});
    if(updated.desc!==undefined&&updated.desc!==(u.desc||""))changes.push({f:"Description",from:u.desc||"—",to:updated.desc||"—"});
    if(updated.projectAccess!==undefined&&updated.projectAccess!==(u.projectAccess||(u.assignedProjectId||toArr(u.assignedProjectIds).length?"specific":"none")))changes.push({f:"Project access",from:u.projectAccess==="all"?"All projects":(u.assignedProjectId||toArr(u.assignedProjectIds).length?"Specific":"None"),to:updated.projectAccess==="all"?"All projects":updated.projectAccess==="specific"?"Specific":"None"});
    if(updated.pin)changes.push({f:"PIN",from:"\u2022\u2022\u2022\u2022",to:"new PIN set"});
    if(updated.phone!==undefined&&updated.phone!==(u.phone||""))changes.push({f:"WhatsApp",from:u.phone||"\u2014",to:updated.phone||"\u2014"});
    if(updated.assignedProjectId!==undefined&&updated.assignedProjectId!==(u.assignedProjectId||""))changes.push({f:"Project",from:pName(u.assignedProjectId),to:pName(updated.assignedProjectId)});
    if(updated.assignedProjectIds!==undefined){
      const a=toArr(u.assignedProjectIds).slice().sort().join(","),b=toArr(updated.assignedProjectIds).slice().sort().join(",");
      if(a!==b)changes.push({f:"Projects",from:toArr(u.assignedProjectIds).map(pName).join(", ")||"None",to:toArr(updated.assignedProjectIds).map(pName).join(", ")||"None"});
    }
    if(updated.caps){
      const oc=u.caps||ROLE_CAPS[roleKey(u.role)]||ROLE_CAPS.engineer;
      PERMS.forEach(([k,l])=>{if(!!updated.caps[k]!==!!oc[k])changes.push({f:l.replace(/[^\w\s]/g,"").trim(),from:oc[k]?"Allowed":"Denied",to:updated.caps[k]?"Allowed":"Denied"});});
    }
    setEditDiff({u,updated,changes});
  }
  function applyEdit(){
    const {u,updated,changes}=editDiff;
    rtdbPatch('users/'+u.id,updated)
      .then(()=>{
        if(updated.assignedProjectId!==u.assignedProjectId) reassignUser({...u,...updated},updated.assignedProjectId);
        logUserAudit(changes.some(c=>c.f==="Project"||c.f==="Projects")?"transferred":"updated",u.name,changes.map(c=>c.f+": "+c.from+" \u2192 "+c.to).join("; "));
        flash("✅ "+u.name+" updated");setEditId(null);
      }).catch(e=>flash(e.message,"err"));
    setEditDiff(null);
  }

  function removeUser(u){
    setDelUser(u);
    setDelCode(""+Math.floor(1000+Math.random()*9000));
    setDelInput("");
  }
  function confirmDelete(){
    if(!delUser)return;
    if(delInput!==delCode){flash("Code does not match","err");return;}
    const u=delUser;
    rtdbDelete('users/'+u.id).then(()=>{logUserAudit("deleted",u.name,"Role: "+(u.role||""));flash("\uD83D\uDDD1 "+u.name+" removed");}).catch(e=>flash(e.message,"err"));
    setDelUser(null);
  }
  function tryCreate(){
    if(!newUsr.name.trim()){flash("Please enter a name first","err");return;}
    setShowCreatePreview(true);
  }

  function toggleMgmtProject(u,pid,add){
    const cur=toArr(u.assignedProjectIds);
    const next=add?[...cur.filter(x=>x!==pid),pid]:cur.filter(x=>x!==pid);
    if(add&&next.length>1)flash("⚠️ "+u.name+" is now assigned to "+next.length+" projects","ok");
    rtdbPatch('users/'+u.id,{assignedProjectIds:next}).then(()=>{logUserAudit("transferred",u.name,"Projects: "+(next.map(pid=>projects.find(p=>p.id===pid)?.name||pid).join(", ")||"None"));flash("Project list updated");}).catch(e=>flash(e.message,"err"));
  }

  const projName=u=>{if(u.projectAccess==="all")return"All projects";const s=projects.find(p=>p.id===u.assignedProjectId);const m=toArr(u.assignedProjectIds).map(pid=>projects.find(p=>p.id===pid)).filter(Boolean);return s?s.name:m.length?m.map(p=>p.name).join(", "):"";};
  const q=search.trim().toLowerCase();
  const filtered=users.filter(u=>{
    if(q&&!((u.name||"").toLowerCase().includes(q)||(u.role||"").toLowerCase().includes(q)||projName(u).toLowerCase().includes(q)))return false;
    if(roleFilter!=="all"&&roleKey(u.role)!==roleKey(roleFilter))return false;
    if(projFilter!=="all"){
      if(projFilter==="unassigned"){if(u.projectAccess==="all"||u.assignedProjectId||toArr(u.assignedProjectIds).length)return false;}
      else{if(u.projectAccess!=="all"&&u.assignedProjectId!==projFilter&&!toArr(u.assignedProjectIds).includes(projFilter))return false;}
    }
    return true;
  });
  let groups=[];
  if(groupBy==="none")groups=[{key:"all",title:"",show:false,users:filtered}];
  else if(groupBy==="role"){
    ["admin","management","incharge","engineer"].forEach(r=>{const us=filtered.filter(u=>roleKey(u.role)===r);if(us.length)groups.push({key:r,title:ROLE_LABELS[roleKey(r)]||r,show:true,users:us});});
    const known=["admin","management","incharge","engineer"];const other=filtered.filter(u=>!known.includes(roleKey(u.role)));if(other.length)groups.push({key:"other",title:"Other",show:true,users:other});
  }else{
    projects.forEach(p=>{const us=filtered.filter(u=>u.projectAccess==="all"||u.assignedProjectId===p.id||toArr(u.assignedProjectIds).includes(p.id));if(us.length)groups.push({key:p.id,title:p.name,show:true,users:us});});
    const un=filtered.filter(u=>u.projectAccess!=="all"&&!u.assignedProjectId&&!toArr(u.assignedProjectIds).length);if(un.length)groups.push({key:"un",title:"Unassigned",show:true,users:un});
  }
  return(
    <div style={{padding:mobile?"8px":"20px",maxWidth:"960px",margin:"0 auto"}}>
      {/* Create User */}
      <Card style={{marginBottom:"14px",borderTop:`4px solid ${NV}`}}>
        <div style={{...SH2,marginBottom:"16px"}}><span>Create User Account</span></div>
        <Grid cols={mobile?"1fr":"1fr 1fr 1fr"}>
          <F lbl="Full Name *"><Inp value={newUsr.name} onChange={e=>setNewUsr(p=>({...p,name:e.target.value}))} placeholder="Name"/></F>
          <F lbl="Role">
            <select value={newUsr.role} onChange={v=>{const r=v.target.value;setNewUsr(p=>({...p,role:r,caps:ROLE_CAPS[roleKey(r)]||ROLE_CAPS.engineer}));}} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px",background:"#fff"}}>
              {roleOpts.map(r=><option key={r}>{r}</option>)}
            </select>
          </F>
          <F lbl="PIN (4+ digits)"><Inp type="password" value={newUsr.pin} onChange={e=>setNewUsr(p=>({...p,pin:e.target.value}))} placeholder="Set a PIN"/></F>
          {roleKey(newUsr.role)==="incharge"&&(
            <F lbl="WhatsApp Number (for notifications)"><Inp type="tel" value={newUsr.phone||""} onChange={e=>setNewUsr(p=>({...p,phone:e.target.value}))} placeholder="e.g. 919876543210 (with country code)"/></F>
          )}
        </Grid>
        <div style={{marginTop:"12px"}}>
          <F lbl="Description"><Inp value={newUsr.desc||""} onChange={e=>setNewUsr(p=>({...p,desc:e.target.value}))} placeholder="Role context, division, notes about this user…"/></F>
        </div>
        {/* Project Access — None / Specific / All projects */}
        <div style={{marginTop:"12px",padding:"12px",background:"#f8fafc",borderRadius:"8px",border:"1px solid #e5e7eb"}}>
          <div style={{fontWeight:"700",fontSize:"12px",color:"#374151",marginBottom:"8px",textTransform:"uppercase",letterSpacing:"0.05em"}}>Project Access</div>
          <div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>
            {[["none","None"],["specific","Specific"],["all","All projects"]].map(([v,l])=>{
              const on=(newUsr.projectAccess||"none")===v;
              return(<button key={v} onClick={()=>setNewUsr(p=>({...p,projectAccess:v,...(v!=="specific"?{assignedProjectId:"",assignedProjectIds:[]}:{})}))} style={{padding:"9px 18px",borderRadius:"8px",border:`2px solid ${on?NV:"#e5e7eb"}`,background:on?"#eff6ff":"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700",color:on?NV:"#6b7280"}}>{l}</button>);
            })}
          </div>
          {(newUsr.projectAccess||"none")==="none"&&(
            <div style={{marginTop:"10px",fontSize:"12px",color:"#92400e",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:"7px",padding:"9px 12px"}}>⚠ This user will be Unassigned and cannot access any project until you assign one.</div>
          )}
          {newUsr.projectAccess==="all"&&(
            <div style={{marginTop:"10px",fontSize:"12px",color:"#166534",background:"#f0fdf4",border:"1px solid #86efac",borderRadius:"7px",padding:"9px 12px"}}>✅ This user will have access to all current and future projects.</div>
          )}
          {newUsr.projectAccess==="specific"&&((roleKey(newUsr.role)==="engineer"||roleKey(newUsr.role)==="incharge")?(
            <div style={{marginTop:"10px"}}>
              <div style={{fontWeight:"700",fontSize:"12px",color:"#0369a1",marginBottom:"8px"}}>📍 Assign to Project <span style={{fontWeight:"400"}}>(one project only)</span></div>
              <select value={newUsr.assignedProjectId||""} onChange={e=>setNewUsr(p=>({...p,assignedProjectId:e.target.value}))} style={{width:"100%",boxSizing:"border-box",padding:"10px 12px",borderRadius:"8px",border:"1.5px solid #7dd3fc",fontSize:"14px",background:"#fff"}}>
                <option value="">— No project yet —</option>
                {projects.map(p=><option key={p.id} value={p.id}>{p.name}{p.code?" ("+p.code+")":""}</option>)}
              </select>
            </div>
          ):(
            <div style={{marginTop:"10px"}}>
              <div style={{fontWeight:"700",fontSize:"12px",color:"#166534",marginBottom:"8px"}}>📍 Assign to Projects <span style={{fontWeight:"400"}}>(can have multiple)</span></div>
              <div style={{display:"flex",flexWrap:"wrap",gap:"8px"}}>
                {projects.map(p=>{
                  const checked=(newUsr.assignedProjectIds||[]).includes(p.id);
                  return(<label key={p.id} style={{display:"flex",alignItems:"center",gap:"6px",padding:"7px 12px",borderRadius:"7px",border:`1.5px solid ${checked?"#22c55e":"#d1d5db"}`,background:checked?"#f0fdf4":"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"600"}}>
                    <input type="checkbox" checked={checked} onChange={e=>{const cur=newUsr.assignedProjectIds||[];const next=e.target.checked?[...cur,p.id]:cur.filter(x=>x!==p.id);setNewUsr(pp=>({...pp,assignedProjectIds:next,assignedProjectId:next[0]||""}));}} style={{accentColor:"#22c55e"}}/>
                    {p.name}
                  </label>);
                })}
              </div>
            </div>
          ))}
        </div>
        {/* Permissions */}
        <div style={{marginTop:"12px",padding:"12px",background:"#f8fafc",borderRadius:"8px",border:"1px solid #e5e7eb"}}>
          <div style={{fontWeight:"700",fontSize:"12px",color:"#374151",marginBottom:"8px",textTransform:"uppercase",letterSpacing:"0.05em"}}>Permissions</div>
          <div style={{display:"grid",gridTemplateColumns:mobile?"1fr 1fr":"repeat(5,1fr)",gap:"7px"}}>
            {PERMS.map(([k,l])=>(
              <label key={k} style={{display:"flex",alignItems:"center",gap:"7px",padding:"9px 10px",borderRadius:"7px",border:`2px solid ${newUsr.caps?.[k]?"#3b82f6":"#e5e7eb"}`,background:newUsr.caps?.[k]?"#eff6ff":"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"600",color:newUsr.caps?.[k]?"#1d4ed8":"#374151"}}>
                <input type="checkbox" checked={!!newUsr.caps?.[k]} onChange={e=>setNewUsr(p=>({...p,caps:{...p.caps,[k]:e.target.checked}}))} style={{width:"14px",height:"14px",accentColor:"#3b82f6"}}/>
                {l}
              </label>
            ))}
          </div>
        </div>
        <div style={{display:"flex",gap:"10px",marginTop:"14px"}}>
          <button onClick={()=>setNewUsr({name:"",role:"engineer",pin:"",desc:"",projectAccess:"none",assignedProjectId:"",assignedProjectIds:[],caps:{fill:true,approve:false,download:false,manage:false,settings:false}})} style={{padding:"13px 20px",borderRadius:"10px",border:"1.5px solid #d1d5db",background:"#fff",color:"#6b7280",cursor:"pointer",fontSize:"14px",fontWeight:"700"}}>Cancel</button>
          <button onClick={tryCreate} style={{flex:1,padding:"13px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"15px",fontWeight:"800",display:"flex",alignItems:"center",justifyContent:"center",gap:"7px"}}><i className="ti ti-eye" aria-hidden/>Preview &amp; confirm</button>
        </div>
      </Card>

      {/* All Users */}
      <Card>
        <div style={{...SH2,marginBottom:"14px"}}><span>All Users ({filtered.length}{filtered.length!==users.length?" of "+users.length:""})</span><button onClick={downloadUsersPDF} style={{padding:"6px 14px",borderRadius:"7px",border:"none",background:"#dc2626",color:"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"700",display:"flex",alignItems:"center",gap:"5px"}}><i className="ti ti-file-type-pdf" aria-hidden/>Download PDF</button></div>
        {/* Toolbar: search, filters, group-by */}
        <div style={{display:"flex",gap:"10px",alignItems:"center",flexWrap:"wrap",marginBottom:"16px"}}>
          <div style={{position:"relative",flex:1,minWidth:"200px"}}>
            <i className="ti ti-search" style={{position:"absolute",left:"12px",top:"50%",transform:"translateY(-50%)",color:"#94a3b8",fontSize:"15px"}} aria-hidden/>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search by name, role or project\u2026" style={{width:"100%",boxSizing:"border-box",padding:"10px 12px 10px 34px",borderRadius:"9px",border:"1.5px solid #d1d5db",fontSize:"14px"}}/>
          </div>
          <select value={roleFilter} onChange={e=>setRoleFilter(e.target.value)} style={{padding:"10px 12px",borderRadius:"9px",border:"1.5px solid #d1d5db",fontSize:"14px",background:"#fff",fontWeight:"600",color:"#334155"}}>
            <option value="all">All roles</option>
            {roleOpts.map(r=><option key={r} value={r}>{ROLE_LABELS[roleKey(r)]||r}</option>)}
          </select>
          <select value={projFilter} onChange={e=>setProjFilter(e.target.value)} style={{padding:"10px 12px",borderRadius:"9px",border:"1.5px solid #d1d5db",fontSize:"14px",background:"#fff",fontWeight:"600",color:"#334155",maxWidth:"180px"}}>
            <option value="all">All projects</option>
            {projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
            <option value="unassigned">Unassigned</option>
          </select>
          <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
            <span style={{fontSize:"12px",color:"#94a3b8",fontWeight:"700"}}>Group</span>
            <div style={{display:"flex",border:"1.5px solid #d1d5db",borderRadius:"8px",overflow:"hidden",background:"#fff"}}>
              {[["none","None"],["role","Role"],["project","Project"]].map(([v,l])=>(
                <button key={v} onClick={()=>setGroupBy(v)} style={{padding:"8px 13px",border:"none",cursor:"pointer",fontSize:"12px",fontWeight:"700",background:groupBy===v?NV:"transparent",color:groupBy===v?"#fff":"#64748b"}}>{l}</button>
              ))}
            </div>
          </div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:"18px"}}>
          {groups.map(g=>(
            <div key={g.key}>
              {g.show&&(<div style={{fontSize:"12px",fontWeight:"800",color:NV,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:"9px",display:"flex",alignItems:"center",gap:"8px"}}>{g.title}<span style={{background:"#eef2f7",color:"#64748b",padding:"1px 9px",borderRadius:"20px",fontSize:"11px"}}>{g.users.length}</span></div>)}
              <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
          {g.users.map(u=>{
            const uc=u.caps||ROLE_CAPS[roleKey(u.role)]||ROLE_CAPS.engineer;
            const isEditing=editId===u.id;
            const ed=isEditing?editData:{};
            const assignedProj=projects.find(p=>p.id===u.assignedProjectId);
            const mgmtProjs=toArr(u.assignedProjectIds).map(pid=>projects.find(p=>p.id===pid)).filter(Boolean);
            return(
              <div key={u.id} style={{border:`2px solid ${isEditing?AM:"#e5e7eb"}`,borderRadius:"12px",overflow:"hidden"}}>
                {/* Header row */}
                <div style={{display:"flex",alignItems:"center",gap:"10px",padding:"12px 14px",background:isEditing?"#fffbeb":"#f8fafc",flexWrap:"wrap"}}>
                  <Av name={u.name} sz={36}/>
                  <div style={{flex:1,minWidth:"100px"}}>
                    {isEditing?(
                      <Inp value={ed.name||u.name} onChange={e=>setEditData(p=>({...p,name:e.target.value}))} style={{fontWeight:"700",fontSize:"14px"}}/>
                    ):(
                      <div style={{fontWeight:"700",fontSize:"14px"}}>{u.name}</div>
                    )}
                    <div style={{display:"flex",gap:"6px",marginTop:"3px",flexWrap:"wrap",alignItems:"center"}}>
                      {isEditing?(
                        <select value={ed.role||u.role} onChange={e=>setEditData(p=>({...p,role:e.target.value,caps:ROLE_CAPS[roleKey(e.target.value)]||ROLE_CAPS.engineer}))} style={{padding:"4px 8px",borderRadius:"6px",border:"1px solid #d1d5db",fontSize:"12px",background:"#fff"}}>
                          {roleOpts.map(r=><option key={r}>{r}</option>)}
                        </select>
                      ):<RoleB role={u.role}/>}
                      {!isEditing&&u.projectAccess==="all"&&<span style={{fontSize:"11px",color:"#166534",background:"#f0fdf4",padding:"2px 8px",borderRadius:"6px",fontWeight:"700"}}>🌐 All projects</span>}
                      {!isEditing&&u.projectAccess!=="all"&&assignedProj&&<span style={{fontSize:"11px",color:"#6b7280",background:"#f3f4f6",padding:"2px 8px",borderRadius:"6px"}}>📍 {assignedProj.name}</span>}
                      {!isEditing&&u.projectAccess!=="all"&&mgmtProjs.length>0&&mgmtProjs.map(p=><span key={p.id} style={{fontSize:"11px",color:"#166534",background:"#f0fdf4",padding:"2px 8px",borderRadius:"6px"}}>📍 {p.name}</span>)}
                      {!isEditing&&u.projectAccess!=="all"&&(roleKey(u.role)==="engineer"||roleKey(u.role)==="incharge")&&!u.assignedProjectId&&<span style={{fontSize:"11px",color:"#d97706",background:"#fffbeb",padding:"2px 8px",borderRadius:"6px"}}>⚠ No project</span>}
                    </div>
                    {!isEditing&&u.desc&&<div style={{fontSize:"12px",color:"#6b7280",marginTop:"4px"}}>{u.desc}</div>}
                  </div>
                  {!isEditing&&<button onClick={()=>startEdit(u)} style={{padding:"7px 14px",borderRadius:"7px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"600",color:NV}}>✏️ Edit</button>}
                  {isEditing&&<>
                    <button onClick={()=>saveEdit(u)} style={{padding:"7px 14px",borderRadius:"7px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"700"}}>👁 Preview changes</button>
                    <button onClick={()=>setEditId(null)} style={{padding:"7px 10px",borderRadius:"7px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"12px",color:"#6b7280"}}>✕</button>
                  </>}
                  <button onClick={()=>removeUser(u)} style={{padding:"7px 12px",borderRadius:"7px",border:"1px solid #fca5a5",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"12px",fontWeight:"600"}}>🗑</button>
                </div>

                {/* Edit body */}
                {isEditing&&(
                  <div style={{padding:"12px 14px",borderTop:"1px solid #fde68a"}}>
                    {/* PIN change */}
                    <div style={{marginBottom:"12px"}}>
                      <div style={{fontSize:"12px",fontWeight:"700",color:"#374151",marginBottom:"6px"}}>Change PIN <span style={{fontWeight:"400",color:"#9ca3af"}}>(leave blank to keep current)</span></div>
                      <Inp type="password" value={ed.pin||""} onChange={e=>setEditData(p=>({...p,pin:e.target.value}))} placeholder="New PIN (4+ digits)" style={{maxWidth:"200px"}}/>
                    </div>
                    {/* Description */}
                    <div style={{marginBottom:"12px"}}>
                      <div style={{fontSize:"12px",fontWeight:"700",color:"#374151",marginBottom:"6px"}}>Description</div>
                      <Inp value={ed.desc!==undefined?ed.desc:u.desc||""} onChange={e=>setEditData(p=>({...p,desc:e.target.value}))} placeholder="Role context, division, notes about this user…"/>
                    </div>
                    {/* Project access mode */}
                    <div style={{marginBottom:"12px"}}>
                      <div style={{fontSize:"12px",fontWeight:"700",color:"#374151",marginBottom:"6px"}}>Project Access</div>
                      <div style={{display:"flex",gap:"7px",flexWrap:"wrap"}}>
                        {[["none","None"],["specific","Specific"],["all","All projects"]].map(([v,l])=>{
                          const on=(ed.projectAccess||"none")===v;
                          return(<button key={v} onClick={()=>setEditData(p=>({...p,projectAccess:v,...(v!=="specific"?{assignedProjectId:"",assignedProjectIds:[]}:{})}))} style={{padding:"7px 15px",borderRadius:"7px",border:`2px solid ${on?NV:"#e5e7eb"}`,background:on?"#eff6ff":"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"700",color:on?NV:"#6b7280"}}>{l}</button>);
                        })}
                      </div>
                      {(ed.projectAccess||"none")==="none"&&<div style={{marginTop:"8px",fontSize:"11px",color:"#92400e",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:"6px",padding:"7px 10px"}}>⚠ This user will be Unassigned and cannot access any project until you assign one.</div>}
                      {ed.projectAccess==="all"&&<div style={{marginTop:"8px",fontSize:"11px",color:"#166534",background:"#f0fdf4",border:"1px solid #86efac",borderRadius:"6px",padding:"7px 10px"}}>✅ This user will have access to all current and future projects.</div>}
                    </div>
                    {roleKey(ed.role||u.role)==="incharge"&&(
                      <div style={{marginBottom:"12px"}}>
                        <div style={{fontSize:"12px",fontWeight:"700",color:"#374151",marginBottom:"6px"}}>WhatsApp Number <span style={{fontWeight:"400",color:"#9ca3af"}}>(for DPR notifications)</span></div>
                        <Inp type="tel" value={ed.phone!==undefined?ed.phone:u.phone||""} onChange={e=>setEditData(p=>({...p,phone:e.target.value}))} placeholder="919876543210" style={{maxWidth:"220px"}}/>
                      </div>
                    )}
                    {/* Project assignment in edit mode — only for Specific access */}
                    {ed.projectAccess!=="specific"?null:(roleKeyroleKey(ed.role||u.role)==="engineer"||roleKeyroleKey(ed.role||u.role)==="incharge")?(
                      <div style={{marginBottom:"12px"}}>
                        <div style={{fontSize:"12px",fontWeight:"700",color:"#374151",marginBottom:"6px"}}>Assigned Project</div>
                        <select value={ed.assignedProjectId||""} onChange={e=>setEditData(p=>({...p,assignedProjectId:e.target.value}))} style={{width:"100%",boxSizing:"border-box",padding:"9px 12px",borderRadius:"7px",border:"1.5px solid #d1d5db",fontSize:"14px",background:"#fff"}}>
                          <option value="">— No project —</option>
                          {projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
                        </select>
                      </div>
                    ):(
                      <div style={{marginBottom:"12px"}}>
                        <div style={{fontSize:"12px",fontWeight:"700",color:"#374151",marginBottom:"6px"}}>Assigned Projects <span style={{fontWeight:"400",color:"#9ca3af"}}>(can have multiple)</span></div>
                        <div style={{display:"flex",flexWrap:"wrap",gap:"6px"}}>
                          {projects.map(p=>{
                            const checked=toArr(ed.assignedProjectIds||u.assignedProjectIds).includes(p.id);
                            return(<label key={p.id} style={{display:"flex",alignItems:"center",gap:"6px",padding:"6px 10px",borderRadius:"6px",border:`1.5px solid ${checked?"#22c55e":"#d1d5db"}`,background:checked?"#f0fdf4":"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"600"}}>
                              <input type="checkbox" checked={checked} onChange={e=>{const cur=toArr(ed.assignedProjectIds||u.assignedProjectIds);const next=e.target.checked?[...cur,p.id]:cur.filter(x=>x!==p.id);setEditData(pp=>({...pp,assignedProjectIds:next,assignedProjectId:next[0]||""}));if(e.target.checked&&next.length>1)flash("⚠️ Assigning to "+next.length+" projects","ok");}} style={{accentColor:"#22c55e"}}/>
                              {p.name}
                            </label>);
                          })}
                        </div>
                      </div>
                    )}
                    {/* Permissions */}
                    <div style={{fontSize:"12px",fontWeight:"700",color:"#374151",marginBottom:"6px"}}>Permissions</div>
                    <div style={{display:"grid",gridTemplateColumns:mobile?"1fr 1fr":"repeat(5,1fr)",gap:"6px"}}>
                      {PERMS.map(([k,l])=>{
                        const checked=!!(ed.caps||uc)[k];
                        return(<label key={k} style={{display:"flex",alignItems:"center",gap:"6px",padding:"7px 10px",borderRadius:"7px",border:`1.5px solid ${checked?"#3b82f6":"#e5e7eb"}`,background:checked?"#eff6ff":"#fff",cursor:roleKey(ed.role||u.role)==="admin"?"default":"pointer",fontSize:"12px",fontWeight:"600",color:checked?"#1d4ed8":"#9ca3af"}}>
                          <input type="checkbox" checked={checked} disabled={roleKey(ed.role||u.role)==="admin"} onChange={e=>setEditData(p=>({...p,caps:{...(p.caps||uc),[k]:e.target.checked}}))} style={{width:"13px",height:"13px",accentColor:"#3b82f6"}}/>
                          {l}
                        </label>);
                      })}
                    </div>
                  </div>
                )}

                {/* Read-only permissions row when not editing */}
                {!isEditing&&(
                  <div style={{padding:"10px 14px",borderTop:"1px solid #f3f4f6",display:"flex",gap:"6px",flexWrap:"wrap"}}>
                    {PERMS.map(([k,l])=>(
                      <span key={k} style={{fontSize:"11px",padding:"3px 9px",borderRadius:"6px",background:uc[k]?"#eff6ff":"#f3f4f6",color:uc[k]?"#1d4ed8":"#9ca3af",fontWeight:"600"}}>{l}</span>
                    ))}
                    {roleKey(u.role)==="admin"&&<span style={{fontSize:"11px",color:"#9ca3af"}}>Admin has all permissions.</span>}
                    {u.createdAt&&<span style={{fontSize:"11px",color:"#9ca3af",marginLeft:"auto"}}>🕓 Created {fmtTs(u.createdAt)}</span>}
                  </div>
                )}
              </div>
            );
          })}
              </div>
            </div>
          ))}
          {filtered.length===0&&(<div style={{textAlign:"center",padding:"44px 20px",color:"#9ca3af",fontSize:"14px"}}><i className="ti ti-user-search" style={{fontSize:"28px",display:"block",marginBottom:"8px"}} aria-hidden/>No users match your filters.</div>)}
        </div>
      </Card>

      {/* User Activity Log */}
      <Card style={{marginTop:"14px"}}>
        <button onClick={()=>setShowLog(s=>!s)} style={{width:"100%",display:"flex",alignItems:"center",gap:"9px",border:"none",background:"transparent",cursor:"pointer",padding:0,textAlign:"left"}}>
          <i className="ti ti-history" style={{fontSize:"17px",color:NV}} aria-hidden/>
          <span style={{fontWeight:"800",fontSize:"15px",color:NV,flex:1}}>User Activity Log <span style={{fontWeight:"600",color:"#94a3b8",fontSize:"12px"}}>({userAudit.length})</span></span>
          <i className={"ti "+(showLog?"ti-chevron-up":"ti-chevron-down")} style={{color:"#94a3b8"}} aria-hidden/>
        </button>
        {showLog&&(
          <div style={{marginTop:"12px",display:"flex",flexDirection:"column"}}>
            {userAudit.length===0&&<div style={{fontSize:"13px",color:"#9ca3af",padding:"10px 0"}}>No user changes recorded yet. Creations, updates, transfers and deletions will appear here.</div>}
            {userAudit.slice(0,50).map((e,i)=>{
              const col=e.action==="created"?{bg:"#f0fdf4",fg:"#166534"}:e.action==="deleted"?{bg:"#fef2f2",fg:"#991b1b"}:e.action==="transferred"?{bg:"#eff6ff",fg:"#1d4ed8"}:{bg:"#fffbeb",fg:"#92400e"};
              const d=e.ts?new Date(e.ts):null;
              return(
                <div key={e.id||i} style={{display:"flex",gap:"10px",alignItems:"flex-start",padding:"10px 2px",borderBottom:"1px solid #f1f5f9"}}>
                  <span style={{fontSize:"11px",fontWeight:"800",padding:"3px 9px",borderRadius:"20px",background:col.bg,color:col.fg,flexShrink:0,textTransform:"capitalize"}}>{e.action}</span>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:"13px",color:"#111827"}}><strong>{e.target}</strong>{e.detail?" \u2014 "+e.detail:""}</div>
                    <div style={{fontSize:"11px",color:"#94a3b8",marginTop:"2px"}}>by {e.actor}{d?" \u00b7 "+d.toLocaleDateString()+" "+d.toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"}):""}</div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </Card>

      {/* Coded delete confirmation */}
      {delUser&&(
        <div onClick={()=>setDelUser(null)} style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",display:"flex",alignItems:"center",justifyContent:"center",padding:"20px",zIndex:2000}}>
          <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:"14px",width:"100%",maxWidth:"400px",overflow:"hidden",boxShadow:"0 30px 60px -20px rgba(0,0,0,.4)"}}>
            <div style={{padding:"16px 20px",background:"#fef2f2",borderBottom:"1px solid #fecaca",display:"flex",alignItems:"center",gap:"10px"}}>
              <i className="ti ti-alert-triangle" style={{color:RD,fontSize:"20px"}} aria-hidden/>
              <div style={{fontWeight:"800",fontSize:"15px",color:"#991b1b"}}>Delete User</div>
            </div>
            <div style={{padding:"20px"}}>
              <div style={{fontSize:"14px",color:"#374151",marginBottom:"10px"}}>Permanently remove <strong>{delUser.name}</strong>? This cannot be undone.</div>
              <div style={{fontSize:"13px",color:"#6b7280",marginBottom:"12px"}}>Type the confirmation code <strong style={{fontFamily:"monospace",fontSize:"17px",color:RD,letterSpacing:"0.18em"}}>{delCode}</strong> below to confirm.</div>
              <Inp value={delInput} onChange={e=>setDelInput(e.target.value)} placeholder="Enter code" style={{fontFamily:"monospace",letterSpacing:"0.1em",textAlign:"center",fontSize:"18px"}}/>
              <div style={{display:"flex",gap:"10px",marginTop:"16px"}}>
                <button onClick={()=>setDelUser(null)} style={{flex:1,padding:"11px",borderRadius:"9px",border:"1.5px solid #d1d5db",background:"#fff",color:"#475569",fontWeight:"700",cursor:"pointer",fontSize:"14px"}}>Cancel</button>
                <button onClick={confirmDelete} disabled={delInput!==delCode} style={{flex:1,padding:"11px",borderRadius:"9px",border:"none",background:delInput===delCode?RD:"#fca5a5",color:"#fff",fontWeight:"800",cursor:delInput===delCode?"pointer":"default",fontSize:"14px"}}>Delete</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Edit diff — review changes before saving */}
      {editDiff&&(
        <div onClick={()=>setEditDiff(null)} style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",display:"flex",alignItems:"flex-start",justifyContent:"center",padding:"40px 20px",overflow:"auto",zIndex:2000}}>
          <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:"16px",width:"100%",maxWidth:"480px",overflow:"hidden",boxShadow:"0 30px 60px -20px rgba(0,0,0,.4)"}}>
            <div style={{padding:"16px 20px",background:AM,display:"flex",alignItems:"center",gap:"9px"}}>
              <i className="ti ti-list-check" style={{color:"#fff",fontSize:"18px"}} aria-hidden/>
              <div style={{color:"#fff",fontWeight:"800",fontSize:"15px"}}>Review Changes — {editDiff.u.name}</div>
            </div>
            <div style={{padding:"20px"}}>
              {editDiff.changes.length===0?(
                <div style={{fontSize:"14px",color:"#6b7280",textAlign:"center",padding:"14px 0"}}>No changes detected.</div>
              ):(
                <div style={{border:"1px solid #e2e8f0",borderRadius:"12px",overflow:"hidden"}}>
                  {editDiff.changes.map((c,i)=>(
                    <div key={i} style={{display:"grid",gridTemplateColumns:"105px 1fr auto 1fr",gap:"8px",alignItems:"center",padding:"10px 14px",borderBottom:i<editDiff.changes.length-1?"1px solid #f1f5f9":"none",fontSize:"13px"}}>
                      <div style={{fontWeight:"700",color:"#6b7280",fontSize:"12px"}}>{c.f}</div>
                      <div style={{color:"#991b1b",background:"#fef2f2",borderRadius:"6px",padding:"3px 8px",fontWeight:"600",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.from}</div>
                      <i className="ti ti-arrow-right" style={{color:"#94a3b8"}} aria-hidden/>
                      <div style={{color:"#166534",background:"#f0fdf4",borderRadius:"6px",padding:"3px 8px",fontWeight:"700",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.to}</div>
                    </div>
                  ))}
                </div>
              )}
              <div style={{display:"flex",gap:"10px",marginTop:"18px"}}>
                <button onClick={()=>setEditDiff(null)} style={{flex:1,padding:"11px",borderRadius:"9px",border:"1.5px solid #d1d5db",background:"#fff",color:"#475569",fontWeight:"700",cursor:"pointer",fontSize:"14px"}}>Back</button>
                {editDiff.changes.length>0&&<button onClick={applyEdit} style={{flex:1,padding:"11px",borderRadius:"9px",border:"none",background:GN,color:"#fff",fontWeight:"800",cursor:"pointer",fontSize:"14px"}}>Confirm &amp; Save</button>}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Create preview / confirm */}
      {showCreatePreview&&(()=>{
        const uc=newUsr.caps||ROLE_CAPS[roleKey(newUsr.role)]||ROLE_CAPS.engineer;
        const proj=projects.find(p=>p.id===newUsr.assignedProjectId);
        const mgmt=toArr(newUsr.assignedProjectIds).map(pid=>projects.find(p=>p.id===pid)).filter(Boolean);
        const access=newUsr.projectAccess||"none";
        const accessCard=access==="all"?"All projects \u2014 full access":access==="specific"?((mgmt.length?mgmt.map(p=>p.name).join(", "):proj?.name)||"No project selected yet"):"Unassigned \u2014 no access yet";
        const accessDetail=access==="all"?"All projects":access==="specific"?((mgmt.length?mgmt.map(p=>p.name).join(", "):proj?.name)||"None selected"):"Unassigned (no access)";
        const grantedPerms=PERMS.filter(([k])=>uc[k]);
        const nowLabel=fmtTs(new Date().toISOString());
        const tsLine=(<div style={{marginTop:"14px",paddingTop:"12px",borderTop:"1px dashed #e2e8f0",fontSize:"12px",color:"#6b7280"}}>Created timestamp will be recorded as<br/><strong style={{color:"#374151",fontSize:"13px"}}>{nowLabel}</strong></div>);
        return(
        <div onClick={()=>setShowCreatePreview(false)} style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",display:"flex",alignItems:"flex-start",justifyContent:"center",padding:"40px 20px",overflow:"auto",zIndex:2000}}>
          <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:"16px",width:"100%",maxWidth:"460px",overflow:"hidden",boxShadow:"0 30px 60px -20px rgba(0,0,0,.4)"}}>
            <div style={{padding:"16px 20px",background:NV,display:"flex",alignItems:"center",gap:"9px"}}>
              <i className="ti ti-user-check" style={{color:"#fff",fontSize:"18px"}} aria-hidden/>
              <div style={{color:"#fff",fontWeight:"800",fontSize:"15px",flex:1}}>Confirm New User</div>
              <div style={{display:"flex",background:"rgba(255,255,255,.15)",borderRadius:"8px",overflow:"hidden"}}>
                {[["card","Card"],["detail","Detail"]].map(([v,l])=>(
                  <button key={v} onClick={()=>setPreviewTab(v)} style={{padding:"6px 14px",border:"none",cursor:"pointer",fontSize:"12px",fontWeight:"700",background:previewTab===v?"#fff":"transparent",color:previewTab===v?NV:"rgba(255,255,255,.85)"}}>{l}</button>
                ))}
              </div>
            </div>
            <div style={{padding:"20px"}}>
              <div style={{fontSize:"13px",color:"#6b7280",marginBottom:"14px"}}>Please review the profile below before it is created.</div>
              {previewTab==="card"?(
                <div style={{border:"1px solid #e2e8f0",borderRadius:"14px",padding:"18px",background:"#f8fafc"}}>
                  <div style={{display:"flex",gap:"14px",alignItems:"center",marginBottom:"12px"}}>
                    <Av name={newUsr.name} sz={48}/>
                    <div><div style={{fontWeight:"800",fontSize:"16px",color:"#111827"}}>{newUsr.name}</div><div style={{marginTop:"3px"}}><RoleB role={newUsr.role}/></div></div>
                  </div>
                  <div style={{fontSize:"13px",color:newUsr.desc?.trim()?"#374151":"#9ca3af",fontStyle:newUsr.desc?.trim()?"normal":"italic",marginBottom:"10px"}}>{newUsr.desc?.trim()||"No description added"}</div>
                  <div style={{fontSize:"13px",fontWeight:"600",color:access==="none"?"#92400e":"#166534",marginBottom:"12px"}}>\ud83d\udccd {accessCard}</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:"5px"}}>
                    {PERMS.map(([k,l])=>(<span key={k} style={{fontSize:"11px",padding:"3px 9px",borderRadius:"6px",background:uc[k]?"#eff6ff":"#f3f4f6",color:uc[k]?"#1d4ed8":"#9ca3af",fontWeight:"600"}}>{l}</span>))}
                  </div>
                  {tsLine}
                </div>
              ):(
                <div style={{border:"1px solid #e2e8f0",borderRadius:"14px",padding:"18px",background:"#f8fafc"}}>
                  <div style={{display:"grid",gridTemplateColumns:"110px 1fr",gap:"9px 12px",fontSize:"13px"}}>
                    <div style={{color:"#6b7280",fontWeight:"600"}}>Name</div><div style={{color:"#111827",fontWeight:"700"}}>{newUsr.name}</div>
                    <div style={{color:"#6b7280",fontWeight:"600"}}>Role</div><div><RoleB role={newUsr.role}/></div>
                    <div style={{color:"#6b7280",fontWeight:"600"}}>Description</div><div style={{color:newUsr.desc?.trim()?"#374151":"#9ca3af"}}>{newUsr.desc?.trim()||"\u2014"}</div>
                    <div style={{color:"#6b7280",fontWeight:"600"}}>Project access</div><div style={{color:"#374151",fontWeight:"600"}}>{accessDetail}</div>
                    {roleKey(newUsr.role)==="incharge"&&newUsr.phone&&(<><div style={{color:"#6b7280",fontWeight:"600"}}>WhatsApp</div><div style={{color:"#374151",fontWeight:"600"}}>{newUsr.phone}</div></>)}
                    <div style={{color:"#6b7280",fontWeight:"600"}}>Permissions</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:"5px"}}>{grantedPerms.map(([k,l])=><span key={k} style={{fontSize:"11px",padding:"2px 8px",borderRadius:"6px",background:"#eff6ff",color:"#1d4ed8",fontWeight:"700"}}>{l}</span>)}{grantedPerms.length===0&&<span style={{fontSize:"12px",color:"#9ca3af"}}>No permissions</span>}</div>
                    <div style={{color:"#6b7280",fontWeight:"600"}}>PIN</div><div style={{fontFamily:"monospace",fontWeight:"700",color:newUsr.pin?RD:"#9ca3af",letterSpacing:"0.1em"}}>{newUsr.pin?"\u2022\u2022\u2022\u2022 set":"Not set"}</div>
                  </div>
                  {tsLine}
                </div>
              )}
              <div style={{display:"flex",gap:"10px",marginTop:"18px"}}>
                <button onClick={()=>setShowCreatePreview(false)} style={{flex:1,padding:"11px",borderRadius:"9px",border:"1.5px solid #d1d5db",background:"#fff",color:"#475569",fontWeight:"700",cursor:"pointer",fontSize:"14px"}}>Back to edit</button>
                <button onClick={()=>{if(newUsr.name.trim()&&newUsr.pin.trim()){const pn=projects.find(p=>p.id===newUsr.assignedProjectId)?.name;logUserAudit("created",newUsr.name.trim(),"Role: "+newUsr.role+(access==="all"?" \u00b7 All projects":pn?" \u00b7 Project: "+pn:""));}addUser();setShowCreatePreview(false);}} style={{flex:1,padding:"11px",borderRadius:"9px",border:"none",background:GN,color:"#fff",fontWeight:"800",cursor:"pointer",fontSize:"14px"}}>Confirm &amp; Create</button>
              </div>
            </div>
          </div>
        </div>
        );
      })()}
    </div>
  );
}

// ─── PROJECTS SCREEN ─────────────────────────────────────────────────────────
function ProjectsScreen({projects,user,users,onEnter,flash,allSubs,globalLists}){
  const [showNew,setShowNew]=useState(false);
  const [np,setNp]=useState({name:"",code:"",location:"",desc:""});
  const [editId,setEditId]=useState(null);
  const [editData,setEditData]=useState({});
  const [assignId,setAssignId]=useState(null); // project being assigned
  const mobile=useMobile();
  const isAdmin=roleKey(user?.role)==="admin";
  const isMgmt=roleKey(user?.role)==="management"||(user?.caps?.download&&user?.caps?.manage);

  // Filter projects by role - use assignedProjectId (single source of truth)
  const visibleProjects=(isAdmin||roleKey(user?.role)==="management"||user?.projectAccess==="all")?projects:projects.filter(p=>{
    if(!user)return false;
    // Incharge: show project they are assigned to
    if(roleKey(user.role)==="incharge") return user.assignedProjectId===p.id;
    return false;
  });

  function createProject(){
    if(!np.name.trim()){flash("Enter project name","err");return;}
    const id=uid();
    rtdbPut('projects/'+id,{id,name:np.name.trim(),code:np.code.trim(),location:np.location.trim(),desc:np.desc.trim(),createdAt:new Date().toISOString(),assignedUsers:[]})
      .then(()=>{flash("✅ Project created");setNp({name:"",code:"",location:"",desc:""});setShowNew(false);})
      .catch(e=>flash("Failed: "+e.message,"err"));
  }

  function saveEdit(id){
    rtdbPatch('projects/'+id,editData)
      .then(()=>{flash("Saved");setEditId(null);})
      .catch(e=>flash("Failed: "+e.message,"err"));
  }

  const [delProj,setDelProj]=useState(null);
  const [delCode,setDelCode]=useState("");
  const [delInput,setDelInput]=useState("");
  const isSuper=isSuperAdminUser(user);
  function deleteProject(id){
    if(!isSuper){flash("Only the Super Admin can delete projects","err");return;}
    const p=projects.find(x=>x.id===id);if(!p)return;
    const code=""+Math.floor(100000+Math.random()*900000);
    setDelProj(p);setDelCode(code);setDelInput("");
  }
  function confirmDeleteProject(){
    if(!delProj||!isSuper)return;
    if(delInput!==delCode){flash("Code does not match","err");return;}
    const p=delProj;
    rtdbDelete('projects/'+p.id).then(()=>flash("\uD83D\uDDD1 "+p.name+" deleted")).catch(e=>flash(e.message,"err"));
    setDelProj(null);
  }

  function toggleUserAssign(project,u,add){
    // Single source of truth: user.assignedProjectId
    const newPid=add?project.id:"";
    rtdbPatch('users/'+u.id,{assignedProjectId:newPid})
      .then(()=>flash(add?"✅ "+u.name+" assigned to "+project.name:"User removed from "+project.name))
      .catch(e=>flash(e.message,"err"));
  }

  const assignableUsers=users.filter(u=>roleKey(u.role)==="incharge"||roleKey(u.role)==="management");

  return(
    <div style={{padding:mobile?"12px":"20px",maxWidth:"1300px",margin:"0 auto"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"24px",flexWrap:"wrap",gap:"12px"}}>
        <div>
          <div style={{fontWeight:"800",fontSize:"22px",color:NV}}>🏗️ SPL Infrastructure — Projects</div>
          <div style={{fontSize:"13px",color:"#6b7280",marginTop:"3px"}}>Select a project to view its dashboard</div>
        </div>
      {isAdmin&&<button onClick={()=>setShowNew(s=>!s)} style={{padding:"12px 22px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"700",display:"flex",alignItems:"center",gap:"8px"}}>
          <i className="ti ti-plus" aria-hidden/>New Project
        </button>}
      </div>

      {showNew&&<Card style={{marginBottom:"20px",border:`2px solid ${NV}`,borderTop:`4px solid ${NV}`}}>
        <div style={{fontWeight:"700",fontSize:"16px",color:NV,marginBottom:"16px"}}>Create New Project</div>
        <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1fr 1fr",gap:"14px"}}>
          <F lbl="Project Name *"><Inp value={np.name} onChange={e=>setNp(p=>({...p,name:e.target.value}))} placeholder="e.g. NH332A Upgradation"/></F>
          <F lbl="Project Code"><Inp value={np.code} onChange={e=>setNp(p=>({...p,code:e.target.value}))} placeholder="e.g. NH-332A"/></F>
          <F lbl="Location"><Inp value={np.location} onChange={e=>setNp(p=>({...p,location:e.target.value}))} placeholder="e.g. Tamil Nadu"/></F>
          <F lbl="Description"><Inp value={np.desc} onChange={e=>setNp(p=>({...p,desc:e.target.value}))} placeholder="Brief description"/></F>
        </div>
        <div style={{display:"flex",gap:"10px",marginTop:"16px"}}>
          <button onClick={createProject} style={{flex:1,padding:"13px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"15px",fontWeight:"800"}}>✅ Create Project</button>
          <button onClick={()=>setShowNew(false)} style={{padding:"13px 16px",borderRadius:"10px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"14px",color:"#6b7280"}}>Cancel</button>
        </div>
      </Card>}

      {visibleProjects.length===0&&<Card style={{textAlign:"center",padding:"4rem",color:"#9ca3af"}}>
        <i className="ti ti-building-estate" style={{fontSize:"48px",display:"block",marginBottom:"12px"}} aria-hidden/>
        <div style={{fontWeight:"700",fontSize:"16px",marginBottom:"6px"}}>{isAdmin?"No projects yet — create one above":"No projects assigned to you"}</div>
        <div style={{fontSize:"13px"}}>Contact admin to assign you to a project.</div>
      </Card>}

      <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":projects.length===1?"1fr":"repeat(auto-fill,minmax(340px,1fr))",gap:"16px"}}>
        {visibleProjects.map(p=>{
          const isEdit=editId===p.id;
          const projIncharges=users.filter(u=>roleKey(u.role)==="incharge"&&u.assignedProjectId===p.id);
          const projEngineers=users.filter(u=>roleKey(u.role)==="engineer"&&u.assignedProjectId===p.id);
          return(
            <div key={p.id} style={{background:"#fff",borderRadius:"14px",border:`2px solid ${NV}20`,overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,.06)"}}>
              <div style={{background:`linear-gradient(135deg,${NV},#2d5a9e)`,padding:"20px 20px 16px",color:"#fff"}}>
                {isEdit?(
                  <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
                    <Inp value={editData.name||p.name} onChange={e=>setEditData(d=>({...d,name:e.target.value}))} style={{fontWeight:"800",fontSize:"18px",background:"rgba(255,255,255,.2)",border:"1px solid rgba(255,255,255,.4)",color:"#fff",borderRadius:"6px",padding:"6px 10px"}}/>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px"}}>
                      <Inp value={editData.code!==undefined?editData.code:p.code||""} onChange={e=>setEditData(d=>({...d,code:e.target.value}))} placeholder="Code" style={{background:"rgba(255,255,255,.2)",border:"1px solid rgba(255,255,255,.4)",color:"#fff",borderRadius:"6px",padding:"5px 8px",fontSize:"13px"}}/>
                      <Inp value={editData.location!==undefined?editData.location:p.location||""} onChange={e=>setEditData(d=>({...d,location:e.target.value}))} placeholder="Location" style={{background:"rgba(255,255,255,.2)",border:"1px solid rgba(255,255,255,.4)",color:"#fff",borderRadius:"6px",padding:"5px 8px",fontSize:"13px"}}/>
                    </div>
                    <div style={{display:"flex",gap:"8px",marginTop:"4px"}}>
                      <button onClick={()=>saveEdit(p.id)} style={{flex:1,padding:"8px",borderRadius:"7px",border:"none",background:"rgba(255,255,255,.9)",color:NV,cursor:"pointer",fontWeight:"700",fontSize:"13px"}}>Save</button>
                      <button onClick={()=>setEditId(null)} style={{padding:"8px 14px",borderRadius:"7px",border:"1px solid rgba(255,255,255,.5)",background:"transparent",color:"#fff",cursor:"pointer",fontSize:"13px"}}>Cancel</button>
                    </div>
                  </div>
                ):(
                  <>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                      <div style={{fontWeight:"800",fontSize:"20px",lineHeight:1.2}}>{p.name}</div>
                      {isAdmin&&<div style={{display:"flex",gap:"6px"}}>
                        <button onClick={()=>{setEditId(p.id);setEditData({});}} style={{padding:"5px 10px",borderRadius:"6px",border:"1px solid rgba(255,255,255,.4)",background:"transparent",cursor:"pointer",color:"#fff",fontSize:"12px"}}>✏️</button>
                        {isSuper&&<button onClick={()=>deleteProject(p.id)} style={{padding:"5px 10px",borderRadius:"6px",border:"1px solid rgba(255,255,255,.4)",background:"transparent",cursor:"pointer",color:"#fca5a5",fontSize:"12px"}}>🗑️</button>}
                      </div>}
                    </div>
                    {p.code&&<div style={{fontSize:"13px",color:"rgba(255,255,255,.75)",marginTop:"4px"}}>📋 {p.code}</div>}
                    {p.location&&<div style={{fontSize:"13px",color:"rgba(255,255,255,.75)",marginTop:"2px"}}>📍 {p.location}</div>}
                    {p.desc&&<div style={{fontSize:"12px",color:"rgba(255,255,255,.6)",marginTop:"4px"}}>{p.desc}</div>}
                  </>
                )}
              </div>
              <div style={{padding:"16px"}}>
                {/* Incharges — from Users tab */}
                <div style={{marginBottom:"10px"}}>
                  <div style={{fontSize:"11px",fontWeight:"700",color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:"5px"}}>
                    Incharges ({projIncharges.length})
                  </div>
                  <div style={{display:"flex",gap:"6px",flexWrap:"wrap"}}>
                    {projIncharges.length>0
                      ?projIncharges.map(u=><span key={u.id} style={{fontSize:"12px",color:"#1e3a5f",background:"#eff6ff",padding:"3px 10px",borderRadius:"20px",fontWeight:"600"}}>👤 {u.name}</span>)
                      :<span style={{fontSize:"12px",color:"#9ca3af"}}>None — assign in Users tab</span>}
                  </div>
                </div>
                {/* Engineers — from Users tab */}
                <div style={{marginBottom:"14px"}}>
                  <div style={{fontSize:"11px",fontWeight:"700",color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:"5px"}}>
                    Engineers ({projEngineers.length})
                  </div>
                  <div style={{display:"flex",gap:"6px",flexWrap:"wrap"}}>
                    {projEngineers.length>0
                      ?projEngineers.map(u=><span key={u.id} style={{fontSize:"12px",color:"#14532d",background:"#f0fdf4",padding:"3px 10px",borderRadius:"20px",fontWeight:"600"}}>🔧 {u.name}</span>)
                      :<span style={{fontSize:"12px",color:"#9ca3af"}}>None — assign in Users tab</span>}
                  </div>
                </div>
                <button onClick={()=>onEnter(p)} style={{width:"100%",padding:"13px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"15px",fontWeight:"800",display:"flex",alignItems:"center",justifyContent:"center",gap:"8px"}}>
                  Enter Project →
                </button>
              </div>
            </div>
          );
        })}
      </div>

      {/* Coded project deletion (OTP-aware) */}
      {delProj&&(
        <div onClick={()=>setDelProj(null)} style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",display:"flex",alignItems:"center",justifyContent:"center",padding:"20px",zIndex:2000}}>
          <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:"14px",width:"100%",maxWidth:"420px",overflow:"hidden",boxShadow:"0 30px 60px -20px rgba(0,0,0,.4)"}}>
            <div style={{padding:"16px 20px",background:"#fef2f2",borderBottom:"1px solid #fecaca",display:"flex",alignItems:"center",gap:"10px"}}>
              <i className="ti ti-alert-triangle" style={{color:RD,fontSize:"20px"}} aria-hidden/>
              <div style={{fontWeight:"800",fontSize:"15px",color:"#991b1b"}}>Delete Project</div>
            </div>
            <div style={{padding:"20px"}}>
              <div style={{fontSize:"14px",color:"#374151",marginBottom:"10px"}}>Permanently delete <strong>{delProj.name}</strong> and all DPRs, engineers &amp; data inside it? This cannot be undone.</div>
              <div style={{fontSize:"13px",color:"#6b7280",marginBottom:"12px"}}>Type the confirmation code <strong style={{fontFamily:"monospace",fontSize:"17px",color:RD,letterSpacing:"0.18em"}}>{delCode}</strong> below.</div>
              <Inp value={delInput} onChange={e=>setDelInput(e.target.value)} placeholder="Enter code" style={{fontFamily:"monospace",letterSpacing:"0.12em",textAlign:"center",fontSize:"18px"}}/>
              <div style={{display:"flex",gap:"10px",marginTop:"16px"}}>
                <button onClick={()=>setDelProj(null)} style={{flex:1,padding:"11px",borderRadius:"9px",border:"1.5px solid #d1d5db",background:"#fff",color:"#475569",fontWeight:"700",cursor:"pointer",fontSize:"14px"}}>Cancel</button>
                <button onClick={confirmDeleteProject} disabled={delInput!==delCode} style={{flex:1,padding:"11px",borderRadius:"9px",border:"none",background:delInput===delCode?RD:"#fca5a5",color:"#fff",fontWeight:"800",cursor:delInput===delCode?"pointer":"default",fontSize:"14px"}}>Delete</button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── ANALYTICS SCREEN (admin) ────────────────────────────────────────────────
function AnalyticsScreen({projects,users,mobile,flash}){
  const [pdata,setPdata]=useState(null); // {[projectId]: {count,pending,lastDate,byEng:{},byDate:{}}}
  const [loading,setLoading]=useState(true);

  async function load(){
    setLoading(true);
    try{
      const entries=await Promise.all(projects.map(async p=>{
        try{
          const r=await fetch(await authedUrl(RTDB_URL+'/projects/'+p.id+'/submissions.json'));
          if(!r.ok) throw new Error('HTTP '+r.status);
          const obj=await r.json()||{};
          const subs=Object.values(obj).filter(Boolean);
          const byEng={},byDate={};let pending=0,lastDate="";
          subs.forEach(s=>{
            const eng=s.engineer||"—";byEng[eng]=(byEng[eng]||0)+1;
            if(s.date){byDate[s.date]=(byDate[s.date]||0)+1;if(s.date>lastDate)lastDate=s.date;}
            if(!s.approved)pending++;
          });
          return [p.id,{count:subs.length,pending,lastDate,byEng,byDate}];
        }catch(e){ return [p.id,{count:0,pending:0,lastDate:"",byEng:{},byDate:{},error:true}]; }
      }));
      setPdata(Object.fromEntries(entries));
    }catch(e){ flash&&flash("Couldn't load analytics: "+e.message,"err"); }
    setLoading(false);
  }
  useEffect(()=>{ load(); /* eslint-disable-next-line */ },[projects.length]);

  // ── Aggregate ──
  const roleCount=r=>users.filter(u=>roleKey(u.role)===r).length;
  const nEng=roleCount("engineer"),nInc=roleCount("incharge"),nMgmt=roleCount("management"),nAdmin=roleCount("admin");
  const unassigned=users.filter(u=>{
    if(u.projectAccess==="all")return false;
    if(roleKey(u.role)==="management")return toArr(u.assignedProjectIds).length===0&&!u.assignedProjectId;
    if(roleKey(u.role)==="admin")return false;
    return !u.assignedProjectId;
  });
  const projIncharges=p=>users.filter(u=>roleKey(u.role)==="incharge"&&u.assignedProjectId===p.id);
  const projEngineers=p=>users.filter(u=>roleKey(u.role)==="engineer"&&u.assignedProjectId===p.id);

  const pd=pdata||{};
  const totalDPRs=Object.values(pd).reduce((a,x)=>a+(x.count||0),0);
  const totalPending=Object.values(pd).reduce((a,x)=>a+(x.pending||0),0);
  const approvedRate=totalDPRs?Math.round((totalDPRs-totalPending)/totalDPRs*100):0;

  // last-14-day submission trend (aggregate across projects)
  const today=new Date();
  const days=[...Array(14)].map((_,i)=>{const d=new Date(today.getTime()-(13-i)*86400000);return d.toISOString().slice(0,10);});
  const trendCounts=days.map(ds=>Object.values(pd).reduce((a,x)=>a+((x.byDate||{})[ds]||0),0));
  const trendMax=Math.max(1,...trendCounts);
  const trendTotal=trendCounts.reduce((a,b)=>a+b,0);
  const engTotals={};Object.values(pd).forEach(x=>Object.entries(x.byEng||{}).forEach(([e,n])=>{engTotals[e]=(engTotals[e]||0)+n;}));
  const topEngs=Object.entries(engTotals).map(([name,count])=>({name,count})).sort((a,b)=>b.count-a.count).slice(0,8);

  const noIncharge=projects.filter(p=>projIncharges(p).length===0);
  const noEngineers=projects.filter(p=>projEngineers(p).length===0);
  const emptyProjects=projects.filter(p=>(pd[p.id]?.count||0)===0);

  const kpis=[
    {label:"Projects",value:projects.length,icon:"ti-building-estate",bg:"#dbeafe",fg:"#1e40af"},
    {label:"Total Users",value:users.length,icon:"ti-users",bg:"#ede9fe",fg:"#5b21b6"},
    {label:"Engineers",value:nEng,icon:"ti-hard-hat",bg:"#dcfce7",fg:"#166534"},
    {label:"DPRs Submitted",value:loading?"…":totalDPRs,icon:"ti-file-check",bg:"#e0f2fe",fg:"#075985"},
    {label:"Pending Approval",value:loading?"…":totalPending,icon:"ti-clock-pause",bg:"#fef3c7",fg:"#92400e"},
    {label:"Approval Rate",value:loading?"…":approvedRate+"%",icon:"ti-discount-check",bg:"#f0fdf4",fg:"#15803d"},
  ];

  const alerts=[];
  if(unassigned.length)alerts.push({icon:"ti-user-off",text:unassigned.length+" user"+(unassigned.length>1?"s have":" has")+" no project assigned — "+unassigned.slice(0,3).map(u=>u.name).join(", ")+(unassigned.length>3?"…":""),bg:"#fef2f2",fg:"#991b1b"});
  if(totalPending)alerts.push({icon:"ti-inbox",text:totalPending+" DPR"+(totalPending>1?"s are":" is")+" awaiting approval across projects",bg:"#eff6ff",fg:"#1e40af"});
  if(noIncharge.length)alerts.push({icon:"ti-user-question",text:noIncharge.length+" project"+(noIncharge.length>1?"s have":" has")+" no incharge assigned — "+noIncharge.slice(0,3).map(p=>p.name).join(", "),bg:"#fffbeb",fg:"#92400e"});
  if(noEngineers.length)alerts.push({icon:"ti-users-minus",text:noEngineers.length+" project"+(noEngineers.length>1?"s have":" has")+" no engineers — DPRs can't be filled there",bg:"#fef2f2",fg:"#991b1b"});
  if(!loading&&emptyProjects.length)alerts.push({icon:"ti-file-off",text:emptyProjects.length+" project"+(emptyProjects.length>1?"s have":" has")+" zero DPR submissions so far",bg:"#f8fafc",fg:"#475569"});
  if(!alerts.length&&!loading)alerts.push({icon:"ti-circle-check",text:"Everything looks healthy — all users assigned and no items pending.",bg:"#f0fdf4",fg:"#15803d"});

  const roleBars=[["Engineers",nEng,"#16a34a"],["Incharges",nInc,"#2563eb"],["Management",nMgmt,"#7c3aed"],["Admin",nAdmin,"#d97706"]];
  const roleMax=Math.max(1,nEng,nInc,nMgmt,nAdmin);

  const SH={fontWeight:"800",fontSize:"14px",color:NV,marginBottom:"14px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:"8px"};

  return(
    <div style={{padding:mobile?"12px":"20px",maxWidth:"1300px",margin:"0 auto"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"20px",flexWrap:"wrap",gap:"10px"}}>
        <div>
          <div style={{fontWeight:"800",fontSize:mobile?"19px":"22px",color:NV}}>📊 Analytics Overview</div>
          <div style={{fontSize:"13px",color:"#6b7280",marginTop:"3px"}}>Users, projects & DPR activity worth noting</div>
        </div>
        <button onClick={load} disabled={loading} style={{padding:"10px 16px",borderRadius:"10px",border:"1px solid #d1d5db",background:"#fff",cursor:loading?"default":"pointer",fontSize:"13px",fontWeight:"700",color:NV,display:"flex",alignItems:"center",gap:"7px",opacity:loading?.6:1}}>
          <i className={"ti "+(loading?"ti-loader-2":"ti-refresh")} aria-hidden/>{loading?"Loading…":"Refresh"}
        </button>
      </div>

      {/* KPI cards */}
      <div style={{display:"grid",gridTemplateColumns:mobile?"1fr 1fr":"repeat(auto-fit,minmax(170px,1fr))",gap:mobile?"10px":"14px",marginBottom:"18px"}}>
        {kpis.map(k=>(
          <div key={k.label} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"14px",padding:mobile?"13px":"16px",boxShadow:"0 1px 4px rgba(0,0,0,.04)"}}>
            <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"10px"}}>
              <div style={{width:"32px",height:"32px",borderRadius:"9px",background:k.bg,color:k.fg,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><i className={"ti "+k.icon} style={{fontSize:"17px"}} aria-hidden/></div>
              <div style={{fontSize:"10px",color:"#6b7280",fontWeight:"700",textTransform:"uppercase",letterSpacing:".04em"}}>{k.label}</div>
            </div>
            <div style={{fontSize:mobile?"24px":"28px",fontWeight:"800",color:"#0f172a",lineHeight:1}}>{k.value}</div>
          </div>
        ))}
      </div>

      {/* Worth noting */}
      <Card style={{marginBottom:"18px"}}>
        <div style={SH}><span>💡 Worth noting</span></div>
        <div style={{display:"flex",flexDirection:"column",gap:"9px"}}>
          {alerts.map((a,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:"11px",background:a.bg,color:a.fg,borderRadius:"10px",padding:"11px 13px",fontSize:"13px",fontWeight:"600"}}>
              <i className={"ti "+a.icon} style={{fontSize:"18px",flexShrink:0}} aria-hidden/><span>{a.text}</span>
            </div>
          ))}
        </div>
      </Card>

      <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1.4fr 1fr",gap:"14px",marginBottom:"18px"}}>
        {/* Trend */}
        <Card>
          <div style={SH}><span>DPR submissions — last 14 days</span><span style={{fontSize:"12px",color:"#6b7280",fontWeight:"600"}}>{loading?"…":trendTotal+" total"}</span></div>
          <div style={{display:"flex",alignItems:"flex-end",gap:mobile?"3px":"5px",height:"120px"}}>
            {trendCounts.map((n,i)=>(
              <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:"4px",height:"100%",justifyContent:"flex-end",minWidth:0}}>
                <div style={{fontSize:"10px",fontWeight:"700",color:"#6b7280",height:"13px"}}>{n||""}</div>
                <div style={{width:"100%",maxWidth:"24px",borderRadius:"5px 5px 2px 2px",height:Math.max(4,Math.round(n/trendMax*78))+"px",background:n?NV:"#e5e7eb"}}/>
                <div style={{fontSize:"9px",color:"#9ca3af",fontWeight:"600"}}>{new Date(days[i]+'T12:00:00').getDate()}</div>
              </div>
            ))}
          </div>
        </Card>
        {/* Role distribution */}
        <Card>
          <div style={SH}><span>Team composition</span><span style={{fontSize:"12px",color:"#6b7280",fontWeight:"600"}}>{users.length} users</span></div>
          <div style={{display:"flex",flexDirection:"column",gap:"13px"}}>
            {roleBars.map(([lbl,n,c])=>(
              <div key={lbl}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"5px"}}><span style={{fontSize:"13px",fontWeight:"700",color:"#374151"}}>{lbl}</span><span style={{fontSize:"13px",fontWeight:"800",color:c}}>{n}</span></div>
                <div style={{background:"#f1f5f9",borderRadius:"99px",height:"8px",overflow:"hidden"}}><div style={{height:"8px",borderRadius:"99px",background:c,width:Math.round(n/roleMax*100)+"%"}}/></div>
              </div>
            ))}
          </div>
        </Card>
      </div>

      <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1.4fr 1fr",gap:"14px"}}>
        {/* Per-project table */}
        <Card>
          <div style={SH}><span>By project</span></div>
          <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
            {projects.length===0&&<div style={{fontSize:"13px",color:"#9ca3af"}}>No projects yet.</div>}
            {projects.map(p=>{
              const d=pd[p.id]||{};const inc=projIncharges(p).length,eng=projEngineers(p).length;
              return(
                <div key={p.id} style={{border:"1px solid #eef2f7",borderRadius:"10px",padding:"12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",gap:"8px",marginBottom:"7px"}}>
                    <div style={{fontWeight:"800",fontSize:"14px",color:NV,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div>
                    <div style={{fontSize:"13px",fontWeight:"800",color:"#0f172a",flexShrink:0}}>{loading?"…":(d.count||0)+" DPRs"}</div>
                  </div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:"6px"}}>
                    <Pill label={"👷 "+eng+" eng"} bg="#f0fdf4" color="#166534"/>
                    <Pill label={"👤 "+inc+" incharge"} bg="#eff6ff" color="#1e40af"/>
                    {!loading&&(d.pending>0)&&<Pill label={"⏳ "+d.pending+" pending"} bg="#fffbeb" color="#92400e"/>}
                    {!loading&&d.lastDate&&<Pill label={"🕑 last "+d.lastDate} bg="#f8fafc" color="#475569"/>}
                    {!loading&&(d.count||0)===0&&<Pill label="no submissions" bg="#fef2f2" color="#991b1b"/>}
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
        {/* Engineer leaderboard */}
        <Card>
          <div style={SH}><span>Most active engineers</span></div>
          {loading?<div style={{fontSize:"13px",color:"#9ca3af"}}>Loading…</div>:
            topEngs.length===0?<div style={{fontSize:"13px",color:"#9ca3af"}}>No submissions yet.</div>:
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {topEngs.map((e,i)=>(
                <div key={e.name} style={{display:"flex",alignItems:"center",gap:"11px"}}>
                  <div style={{width:"22px",fontSize:"13px",fontWeight:"800",color:i<3?AM:"#cbd5e1",textAlign:"center"}}>{i+1}</div>
                  <Av name={e.name} sz={30}/>
                  <div style={{flex:1,minWidth:0,fontSize:"13px",fontWeight:"700",color:"#374151",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.name}</div>
                  <div style={{fontSize:"13px",fontWeight:"800",color:NV,flexShrink:0}}>{e.count}</div>
                </div>
              ))}
            </div>}
        </Card>
      </div>
    </div>
  );
}

// ─── PERFORMANCE SCREEN (admin) — star ratings leaderboard ──────────────────
function PerformanceScreen({users,projects,mobile,globalLists}){
  const [allSubs,setAllSubs]=useState(null);
  const [loading,setLoading]=useState(true);
  useEffect(()=>{
    let cancelled=false;
    setLoading(true);
    fetchAllSubmissions(projects).then(subs=>{ if(!cancelled){setAllSubs(subs);setLoading(false);} });
    return()=>{cancelled=true;};
  },[projects.length]);

  const starStart=Number(globalLists.starStart??5);
  const subs=allSubs||[];
  const isLateSub=s=>s.date!==((s.submittedAt||"").slice(0,10)||s.date);
  const people=users.filter(u=>roleKey(u.role)==="engineer"||roleKey(u.role)==="incharge").map(u=>{
    const mine=roleKey(u.role)==="engineer"?subs.filter(s=>s.engineer===u.name):subs.filter(s=>s.incharge===u.name);
    const late=mine.filter(isLateSub).length;
    return{ id:u.id, name:u.name, role:u.role, stars:u.stars!=null?Number(u.stars):starStart, subs:mine.length, late, ontime:mine.length-late };
  }).sort((a,b)=>b.stars-a.stars);

  return(
    <div style={{padding:mobile?"12px":"20px",maxWidth:"1100px",margin:"0 auto"}}>
      <div style={{fontWeight:"800",fontSize:mobile?"19px":"22px",color:NV,marginBottom:"3px"}}>⭐ Performance</div>
      <div style={{fontSize:"13px",color:"#6b7280",marginBottom:"16px"}}>Engineer &amp; incharge ratings based on on-time submission</div>
      <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:"10px",padding:"13px 16px",marginBottom:"18px",fontSize:"13px",color:"#92400e",lineHeight:"1.5"}}>
        <i className="ti ti-star-filled" style={{color:"#f59e0b"}} aria-hidden/> Everyone starts at <strong>{starStart}★</strong>. A late submission deducts <strong>{globalLists.lateEngDeduct??0.5}★</strong> from the engineer and <strong>{globalLists.lateInchargeDeduct??0.25}★</strong> from their incharge. Configurable in Settings.
      </div>
      {loading&&<div style={{fontSize:"13px",color:"#9ca3af",padding:"20px 0"}}>Loading submission history…</div>}
      {!loading&&(
        <Card style={{padding:0,overflow:"hidden"}}>
          <div style={{display:"grid",gridTemplateColumns:mobile?"32px 1fr 70px 60px 60px":"40px 1.6fr 1.2fr 1fr 90px 80px",gap:"10px",padding:mobile?"10px 12px":"11px 16px",background:"#f8fafc",borderBottom:"1px solid #e5e7eb",fontSize:"10px",fontWeight:"800",color:"#6b7280",textTransform:"uppercase"}}>
            <div>#</div><div>Name</div>{!mobile&&<div>Role</div>}<div>Stars</div><div style={{textAlign:"center"}}>On-time</div><div style={{textAlign:"center"}}>Late</div>
          </div>
          {people.length===0&&<div style={{padding:"30px",textAlign:"center",color:"#9ca3af",fontSize:"13px"}}>No engineers or incharges yet.</div>}
          {people.map((p,i)=>(
            <div key={p.id} style={{display:"grid",gridTemplateColumns:mobile?"32px 1fr 70px 60px 60px":"40px 1.6fr 1.2fr 1fr 90px 80px",gap:"10px",padding:mobile?"10px 12px":"13px 16px",borderBottom:"1px solid #f1f5f9",alignItems:"center",fontSize:"13px"}}>
              <div style={{fontWeight:"800",color:i===0?AM:"#94a3b8"}}>{i+1}</div>
              <div style={{fontWeight:"700",color:"#0f172a",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div>
              {!mobile&&<div><RoleB role={p.role}/></div>}
              <div style={{fontWeight:"800",color:"#475569",display:"flex",alignItems:"center",gap:"3px"}}><i className="ti ti-star-filled" style={{fontSize:"13px",color:"#f59e0b"}} aria-hidden/>{p.stars.toFixed(2)}</div>
              <div style={{textAlign:"center",fontWeight:"700",color:GN}}>{p.ontime}</div>
              <div style={{textAlign:"center",fontWeight:"700",color:p.late>0?RD:"#cbd5e1"}}>{p.late}</div>
            </div>
          ))}
        </Card>
      )}
    </div>
  );
}

// ─── AUDIT LOG SCREEN (admin) ────────────────────────────────────────────────
function AuditLogScreen({auditTrail,mobile}){
  return(
    <div style={{padding:mobile?"12px":"20px",maxWidth:"1000px",margin:"0 auto"}}>
      <div style={{fontWeight:"800",fontSize:mobile?"19px":"22px",color:NV,marginBottom:"3px"}}>🕘 Audit Log</div>
      <div style={{fontSize:"13px",color:"#6b7280",marginBottom:"16px"}}>Permanent trail of late submissions &amp; backlog approvals</div>
      <div style={{background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:"10px",padding:"13px 16px",marginBottom:"18px",fontSize:"13px",color:"#1e40af",lineHeight:"1.5"}}>
        <i className="ti ti-history" style={{color:"#2563eb"}} aria-hidden/> Every late submission and approved backlog entry is recorded here — engineer, incharge, approver and the star penalty applied — separate from the live Performance scores.
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:"9px"}}>
        {auditTrail.length===0&&<div style={{textAlign:"center",padding:"40px",color:"#9ca3af",fontWeight:"600",fontSize:"13px"}}>No audit entries yet — they'll appear once a late submission or backlog approval occurs.</div>}
        {auditTrail.map(l=>{
          const isBacklog=l.type==="backlog_approved";
          return(
            <div key={l.id} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"11px",padding:"13px 15px",display:"flex",alignItems:"center",gap:"14px",flexWrap:"wrap"}}>
              <i className={"ti "+(isBacklog?"ti-history":"ti-clock-exclamation")} style={{fontSize:"20px",color:isBacklog?"#2563eb":"#d97706",flexShrink:0}} aria-hidden/>
              <div style={{flex:1,minWidth:"160px"}}>
                <div style={{fontSize:"13px",color:"#0f172a",fontWeight:"700"}}>{isBacklog?"Backlog entry approved":"Late submission"} — {l.engName}</div>
                <div style={{fontSize:"12px",color:"#6b7280",marginTop:"2px"}}>{l.detail}{l.projectName?" · "+l.projectName:""}</div>
              </div>
              <div style={{display:"flex",gap:"7px",flexWrap:"wrap"}}>
                {l.engDeduct>0&&<span style={{fontSize:"11px",fontWeight:"800",background:"#fef2f2",color:RD,borderRadius:"7px",padding:"4px 9px"}}>Engineer −{l.engDeduct}★</span>}
                {l.icDeduct>0&&<span style={{fontSize:"11px",fontWeight:"800",background:"#fff7ed",color:"#c2410c",borderRadius:"7px",padding:"4px 9px"}}>Incharge −{l.icDeduct}★</span>}
              </div>
              <div style={{fontSize:"11px",color:"#94a3b8",width:"100%",borderTop:"1px solid #f1f5f9",paddingTop:"8px",marginTop:"2px"}}>
                <i className="ti ti-signature" style={{fontSize:"12px"}} aria-hidden/> {isBacklog?"Approved by":"Detected by"} <strong style={{color:"#475569"}}>{l.approver}</strong> · {l.ts?new Date(l.ts).toLocaleString("en-IN",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"}):""}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── LATE-ENTRY REQUESTS SCREEN (admin) ──────────────────────────────────────
function LateRequestsScreen({lateRequests,decideLateRequest,mobile}){
  const sorted=[...lateRequests].sort((a,b)=>{
    if((a.status==="pending")!==(b.status==="pending"))return a.status==="pending"?-1:1;
    return (b.requestedAt||"").localeCompare(a.requestedAt||"");
  });
  return(
    <div style={{padding:mobile?"12px":"20px",maxWidth:"860px",margin:"0 auto"}}>
      <div style={{fontWeight:"800",fontSize:mobile?"19px":"22px",color:NV,marginBottom:"3px"}}>🕗 Late-Entry Requests</div>
      <div style={{fontSize:"13px",color:"#6b7280",marginBottom:"18px"}}>Approve or reject requests to submit DPRs past the date-lock window</div>
      <div style={{display:"flex",flexDirection:"column",gap:"11px"}}>
        {sorted.length===0&&<div style={{textAlign:"center",padding:"40px",color:"#9ca3af",fontWeight:"600",fontSize:"13px"}}>No late-entry requests.</div>}
        {sorted.map(r=>(
          <Card key={r.id}>
            <div style={{display:"flex",alignItems:"flex-start",gap:"12px",flexWrap:"wrap"}}>
              <Av name={r.engName} sz={38}/>
              <div style={{flex:1,minWidth:"180px"}}>
                <div style={{fontWeight:"800",fontSize:"14px",color:"#0f172a"}}>{r.engName}</div>
                <div style={{fontSize:"13px",color:"#334155",marginTop:"3px"}}>Requesting entry for <strong>{r.date}</strong>{r.projectName?" — "+r.projectName:""}</div>
                <div style={{fontSize:"13px",color:"#64748b",marginTop:"6px",background:"#f8fafc",borderRadius:"8px",padding:"8px 10px",border:"1px solid #f1f5f9"}}><i className="ti ti-message" style={{fontSize:"12px",color:"#94a3b8"}} aria-hidden/> {r.reason||"(no reason given)"}</div>
                <div style={{fontSize:"11px",color:"#94a3b8",marginTop:"6px"}}>Raised {r.requestedAt?new Date(r.requestedAt).toLocaleString("en-IN",{day:"2-digit",month:"short",hour:"2-digit",minute:"2-digit"}):""}</div>
                {r.status!=="pending"&&<div style={{fontSize:"11px",color:"#94a3b8",marginTop:"4px"}}>{r.status==="approved"?"Approved":"Rejected"} by {r.decidedBy} on {r.decidedAt?new Date(r.decidedAt).toLocaleDateString("en-IN"):""}</div>}
              </div>
              {r.status==="pending"?(
                <div style={{display:"flex",flexDirection:mobile?"row":"column",gap:"7px",flexShrink:0}}>
                  <button onClick={()=>decideLateRequest(r.id,true)} style={{display:"inline-flex",alignItems:"center",gap:"5px",padding:"8px 14px",borderRadius:"8px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"800"}}><i className="ti ti-check" aria-hidden/>Approve</button>
                  <button onClick={()=>decideLateRequest(r.id,false)} style={{display:"inline-flex",alignItems:"center",gap:"5px",padding:"8px 14px",borderRadius:"8px",border:"1.5px solid #fecaca",background:"#fff",color:RD,cursor:"pointer",fontSize:"12px",fontWeight:"700"}}><i className="ti ti-x" aria-hidden/>Reject</button>
                </div>
              ):(
                <span style={{flexShrink:0,fontSize:"12px",fontWeight:"800",color:r.status==="approved"?GN:RD,background:r.status==="approved"?"#dcfce7":"#fee2e2",borderRadius:"8px",padding:"6px 12px"}}>
                  <i className={"ti "+(r.status==="approved"?"ti-circle-check":"ti-circle-x")} aria-hidden/> {r.status==="approved"?"Approved":"Rejected"}
                </span>
              )}
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

// ─── REPORTS SCREEN (admin) — consolidated cross-project report ─────────────
function ReportsScreen({projects,users,mobile,flash}){
  const todayStr=new Date().toISOString().slice(0,10);
  const [from,setFrom]=useState(new Date(Date.now()-6*86400000).toISOString().slice(0,10));
  const [to,setTo]=useState(todayStr);
  const [allSubs,setAllSubs]=useState(null);
  const [loading,setLoading]=useState(true);

  useEffect(()=>{
    let cancelled=false;
    setLoading(true);
    fetchAllSubmissions(projects).then(subs=>{ if(!cancelled){setAllSubs(subs);setLoading(false);} });
    return()=>{cancelled=true;};
  },[projects.length]);

  const quickRanges=[
    ["Today",()=>{setFrom(todayStr);setTo(todayStr);}],
    ["Last 7 days",()=>{setFrom(new Date(Date.now()-6*86400000).toISOString().slice(0,10));setTo(todayStr);}],
    ["Last 30 days",()=>{setFrom(new Date(Date.now()-29*86400000).toISOString().slice(0,10));setTo(todayStr);}],
    ["This month",()=>{const d=new Date();setFrom(new Date(d.getFullYear(),d.getMonth(),1).toISOString().slice(0,10));setTo(todayStr);}],
  ];

  const subs=allSubs||[];
  const inRange=subs.filter(s=>s.date>=from&&s.date<=to);
  const engRoster=users.filter(u=>roleKey(u.role)==="engineer");

  const lastFieldFor=(name,field)=>{ const ss=inRange.filter(s=>s.engineer===name); return ss.length?ss[ss.length-1][field]||"":""; };
  const engRows=engRoster.map(u=>{
    const ss=inRange.filter(s=>s.engineer===u.name);
    return{
      name:u.name, dept:lastFieldFor(u.name,"dept")||"—",
      days:new Set(ss.map(s=>s.date)).size, subs:ss.length,
      acts:ss.reduce((a,s)=>a+(s.activities||[]).length,0),
      mats:ss.reduce((a,s)=>a+(s.matTxs||[]).length,0),
      apvd:ss.filter(s=>s.approved).length, pend:ss.filter(s=>!s.approved).length,
    };
  }).filter(r=>r.subs>0).sort((a,b)=>b.subs-a.subs);

  const dayRows=(()=>{
    const rows=[]; let d=new Date(from+"T12:00:00"); const end=new Date(to+"T12:00:00");
    while(d<=end){
      const ds=d.toISOString().slice(0,10);
      const ss=inRange.filter(s=>s.date===ds);
      const present=new Set(ss.map(s=>s.engineer)).size;
      rows.push({date:ds, dow:d.toLocaleDateString("en-IN",{weekday:"short"}), present, absent:Math.max(0,engRoster.length-present), acts:ss.reduce((a,s)=>a+(s.activities||[]).length,0), apvd:ss.filter(s=>s.approved).length, total:ss.length});
      d=new Date(d.getTime()+86400000);
    }
    return rows.reverse();
  })();

  const metrics=[
    {label:"Submissions",value:inRange.length,color:NV,bg:"#eff6ff"},
    {label:"Approved",value:inRange.filter(s=>s.approved).length,color:GN,bg:"#f0fdf4"},
    {label:"Pending",value:inRange.filter(s=>!s.approved).length,color:AM,bg:"#fffbeb"},
    {label:"Active Engineers",value:engRows.length,color:PU,bg:"#f5f3ff"},
  ];

  function download(){
    if(!inRange.length){flash("No data in this range","err");return;}
    const engForExport=engRoster.map(u=>({name:u.name,dept:lastFieldFor(u.name,"dept"),incharge:lastFieldFor(u.name,"incharge")}));
    doExcel(inRange,engForExport,from+" to "+to,"All Projects");
    flash("✅ Report downloaded");
  }

  return(
    <div style={{padding:mobile?"12px":"20px",maxWidth:"1100px",margin:"0 auto"}}>
      <div style={{fontWeight:"800",fontSize:mobile?"19px":"22px",color:NV,marginBottom:"3px"}}>📊 Reports</div>
      <div style={{fontSize:"13px",color:"#6b7280",marginBottom:"16px"}}>Consolidated DPR report across all projects</div>

      <Card style={{marginBottom:"16px",borderTop:`4px solid ${NV}`}}>
        <div style={{display:"grid",gridTemplateColumns:mobile?"1fr 1fr":"1fr 1fr auto",gap:"14px",alignItems:"end"}}>
          <F lbl="From Date"><input type="date" value={from} onChange={e=>setFrom(e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"14px",fontWeight:"600"}}/></F>
          <F lbl="To Date"><input type="date" value={to} onChange={e=>setTo(e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"14px",fontWeight:"600"}}/></F>
          <button onClick={download} style={{display:"inline-flex",alignItems:"center",gap:"6px",padding:"11px 16px",borderRadius:"9px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"800",whiteSpace:"nowrap"}}><i className="ti ti-file-spreadsheet" aria-hidden/>Download</button>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:"8px",marginTop:"14px"}}>
          <span style={{fontSize:"12px",color:"#94a3b8",fontWeight:"700",alignSelf:"center"}}>Quick range:</span>
          {quickRanges.map(([l,fn])=>(
            <button key={l} onClick={fn} style={{padding:"6px 13px",borderRadius:"8px",border:"1.5px solid #e2e8f0",background:"#f8fafc",color:"#475569",cursor:"pointer",fontSize:"12px",fontWeight:"700"}}>{l}</button>
          ))}
        </div>
      </Card>

      {loading&&<div style={{fontSize:"13px",color:"#9ca3af",padding:"20px 0"}}>Loading submission history…</div>}
      {!loading&&inRange.length===0&&<Card style={{textAlign:"center",padding:"50px",color:"#9ca3af"}}><i className="ti ti-calendar-off" style={{fontSize:"32px"}} aria-hidden/><div style={{marginTop:"10px",fontWeight:"600"}}>No DPR data for {from} to {to}</div></Card>}
      {!loading&&inRange.length>0&&(<>
        <div style={{display:"grid",gridTemplateColumns:mobile?"1fr 1fr":"repeat(4,1fr)",gap:"12px",marginBottom:"16px"}}>
          {metrics.map(m=>(
            <div key={m.label} style={{background:m.bg,borderRadius:"12px",padding:"15px 16px"}}>
              <div style={{fontSize:"11px",color:"#64748b",fontWeight:"700",textTransform:"uppercase",letterSpacing:".04em",marginBottom:"6px"}}>{m.label}</div>
              <div style={{fontSize:"28px",fontWeight:"800",color:m.color}}>{m.value}</div>
            </div>
          ))}
        </div>
        <Card style={{padding:0,overflow:"hidden",marginBottom:"16px"}}>
          <div style={{padding:"14px 18px",fontWeight:"800",fontSize:"14px",color:"#0f172a",borderBottom:"1px solid #f1f5f9"}}>Engineer Summary</div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px",minWidth:"560px"}}>
              <thead><tr style={{background:"#f8fafc"}}>{["Engineer","Dept","Days","Subs","Activities","Mat. Moves","Apvd","Pend"].map(h=><th key={h} style={{padding:"10px 14px",textAlign:"left",fontSize:"11px",fontWeight:"800",color:"#64748b",textTransform:"uppercase",borderBottom:"1px solid #e5e7eb",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
              <tbody>
                {engRows.map(e=>(
                  <tr key={e.name} style={{borderBottom:"1px solid #f3f4f6"}}>
                    <td style={{padding:"11px 14px",fontWeight:"700",color:"#0f172a"}}>{e.name}</td>
                    <td style={{padding:"11px 14px",color:"#475569"}}>{e.dept}</td>
                    <td style={{padding:"11px 14px",fontWeight:"700",color:NV}}>{e.days}</td>
                    <td style={{padding:"11px 14px",fontWeight:"700",color:PU}}>{e.subs}</td>
                    <td style={{padding:"11px 14px",fontWeight:"700"}}>{e.acts}</td>
                    <td style={{padding:"11px 14px",fontWeight:"700"}}>{e.mats}</td>
                    <td style={{padding:"11px 14px"}}><Pill label={String(e.apvd)} bg="#f0fdf4" color={GN}/></td>
                    <td style={{padding:"11px 14px"}}>{e.pend>0?<Pill label={String(e.pend)} bg="#fffbeb" color="#92400e"/>:<span style={{color:"#cbd5e1"}}>0</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
        <Card style={{padding:0,overflow:"hidden"}}>
          <div style={{padding:"14px 18px",fontWeight:"800",fontSize:"14px",color:"#0f172a",borderBottom:"1px solid #f1f5f9"}}>Day-wise Attendance</div>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px",minWidth:"520px"}}>
              <thead><tr style={{background:"#f8fafc"}}>{["Date","Day","Present","Absent","Activities","Approved"].map(h=><th key={h} style={{padding:"10px 14px",textAlign:"left",fontSize:"11px",fontWeight:"800",color:"#64748b",textTransform:"uppercase",borderBottom:"1px solid #e5e7eb"}}>{h}</th>)}</tr></thead>
              <tbody>
                {dayRows.map(d=>(
                  <tr key={d.date} style={{borderBottom:"1px solid #f3f4f6"}}>
                    <td style={{padding:"10px 14px",fontWeight:"700",fontFamily:"ui-monospace,monospace",color:"#0f172a"}}>{d.date}</td>
                    <td style={{padding:"10px 14px",color:"#64748b"}}>{d.dow}</td>
                    <td style={{padding:"10px 14px"}}><Pill label={String(d.present)} bg="#f0fdf4" color={GN}/></td>
                    <td style={{padding:"10px 14px"}}>{d.absent>0?<Pill label={String(d.absent)} bg="#fef2f2" color={RD}/>:<span style={{color:"#cbd5e1"}}>0</span>}</td>
                    <td style={{padding:"10px 14px",fontWeight:"700"}}>{d.acts}</td>
                    <td style={{padding:"10px 14px",fontWeight:"800",color:GN}}>{d.apvd}/{d.total}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      </>)}
    </div>
  );
}

// ─── SETTINGS PANEL ──────────────────────────────────────────────────────────
function SettingsPanel({settingsTab,setSettingsTab,listEdits,setListEdits,lists,setLists,workTypes,saveWt,flash,projectId}){
  const TABS=[
    ["workTypes","📋 Work Types",workTypes,DEFAULT_WT],
    ["assetGroups","🚛 Machines",lists.assetGroups,ASSET_GROUPS],
    ["labourTypes","👷 Labour Types",lists.labourTypes,LABOUR_TYPES],
    ["depts","🏗️ Departments",lists.depts,DEPTS],
    ["designations","🏷️ Designations",lists.designations,DESIGNATIONS],
    ["sides","📍 Sides",lists.sides,SIDES],
    ["units","📏 Units",lists.units,UNITS],
    ["bulkMats","🪨 Materials",lists.bulkMats,BULK_MATS],
    ["prodMats","🏭 Prod. Materials",lists.prodMats,PROD_MATS.map(x=>x.n)],
  ];
  const [tab,setTab]=useState(settingsTab||"workTypes");
  const current=TABS.find(t=>t[0]===tab)||TABS[0];
  const [key,label,data,def]=current;
  const dataArr=Array.isArray(data)?data:def;
  const val=listEdits[key]!==undefined?listEdits[key]:dataArr.join("\n");
  const isProdMats=key==="prodMats";

  function save(){
    const lines=val.split("\n").map(s=>s.trim()).filter(Boolean);
    if(!lines.length){flash("List cannot be empty","err");return;}
    if(key==="workTypes"){saveWt(lines);return;}
    if(!projectId)return;
    rtdbPut('projects/'+projectId+'/config/editableLists',{...lists,[key]:lines})
      .then(()=>{setLists(p=>({...p,[key]:lines}));flash("✅ Saved — "+lines.length+" items");})
      .catch(e=>flash("Failed: "+e.message,"err"));
  }

  return(
    <div>
      <div style={{display:"flex",gap:"6px",marginBottom:"16px",flexWrap:"wrap"}}>
        {TABS.map(([k,,d])=>(
          <button key={k} onClick={()=>{setTab(k);setSettingsTab(k);}} style={{padding:"8px 14px",borderRadius:"8px",border:`2px solid ${tab===k?NV:"#e5e7eb"}`,background:tab===k?NV:"#fff",color:tab===k?"#fff":"#374151",cursor:"pointer",fontSize:"13px",fontWeight:tab===k?"700":"500"}}>
            {TABS.find(t=>t[0]===k)?.[1]||k}
          </button>
        ))}
      </div>
      <Card>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"8px"}}>
          <div style={{fontWeight:"700",fontSize:"16px",color:NV}}>{label}</div>
          <span style={{fontSize:"12px",color:"#9ca3af"}}>{dataArr.length} items</span>
        </div>
        {isProdMats&&<p style={{fontSize:"12px",color:"#0369a1",background:"#eff6ff",padding:"8px 12px",borderRadius:"6px",marginBottom:"8px",marginTop:0}}>One material name per line. e.g. <code>M25 RMC</code></p>}
        <p style={{fontSize:"13px",color:"#6b7280",marginBottom:"12px",marginTop:0}}>One entry per line. Saved to Firebase — reflects in all dropdowns immediately.</p>
        <textarea value={val} onChange={e=>setListEdits(p=>({...p,[key]:e.target.value}))} rows={18}
          style={{width:"100%",boxSizing:"border-box",fontFamily:"monospace",fontSize:"14px",padding:"14px",border:"2px solid #d1d5db",borderRadius:"10px",background:"#f9fafb",color:"#111827",lineHeight:"1.9",resize:"vertical"}}/>
        <div style={{display:"flex",gap:"10px",marginTop:"14px"}}>
          <button onClick={save} style={{flex:1,padding:"13px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"15px",fontWeight:"800"}}>💾 Save Changes</button>
          <button onClick={()=>setListEdits(p=>({...p,[key]:def.join("\n")}))} style={{padding:"13px 16px",borderRadius:"10px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"14px",color:"#6b7280"}}>Reset Default</button>
        </div>
      </Card>
    </div>
  );
}

// ─── SHARE MODAL ─────────────────────────────────────────────────────────────
function ShareModal({subs,engineers,dashDate,reportFrom,reportTo,onClose,flash}){
  const [shareType,setShareType]=useState("day");
  const [shareFrom,setShareFrom]=useState(reportFrom||dashDate);
  const [shareTo,setShareTo]=useState(reportTo||dashDate);
  function getShareData(){
    const filtered=shareType==="day"?subs.filter(s=>s.date===shareFrom):subs.filter(s=>s.date>=shareFrom&&s.date<=shareTo);
    const days=[...new Set(filtered.map(s=>s.date))].sort();
    const pres=[...new Set(filtered.map(s=>s.engineer))];
    const acts=filtered.reduce((a,s)=>a+(s.activities||[]).length,0);
    const mats=filtered.reduce((a,s)=>a+(s.matTxs||[]).length,0);
    const apvd=filtered.filter(s=>s.approved).length;
    const period=shareType==="day"?shareFrom:`${shareFrom} to ${shareTo}`;
    let text=`*SPL Infrastructure — NH332A*\n*DPR Summary (${period})*\n\n`;
    text+=`📅 Days: ${days.length}  |  👷 Engineers: ${pres.length}\n`;
    text+=`🔧 Activities: ${acts}  |  📦 Material Moves: ${mats}\n`;
    text+=`✅ Approved: ${apvd}/${filtered.length}\n\n`;
    if(shareType==="day"){
      const present=engineers.filter(e=>filtered.some(s=>s.engineer===e.name));
      const absent=engineers.filter(e=>!filtered.some(s=>s.engineer===e.name));
      if(present.length)text+=`*Present (${present.length}):* ${present.map(e=>e.name).join(", ")}\n`;
      if(absent.length)text+=`*Absent (${absent.length}):* ${absent.map(e=>e.name).join(", ")}\n`;
    }
    return text;
  }
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:99999,display:"flex",alignItems:"center",justifyContent:"center",padding:"16px"}}>
      <div style={{background:"#fff",borderRadius:"16px",padding:"24px",maxWidth:"500px",width:"100%",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{fontWeight:"800",fontSize:"18px",color:NV,marginBottom:"20px"}}>📤 Share Progress Report</div>
        <div style={{marginBottom:"16px"}}><Seg opts={["day","range"]} val={shareType} onChange={setShareType}/></div>
        {shareType==="day"?(
          <F lbl="Select Date"><input type="date" value={shareFrom} onChange={e=>setShareFrom(e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px"}}/></F>
        ):(
          <Grid cols="1fr 1fr">
            <F lbl="From"><input type="date" value={shareFrom} onChange={e=>setShareFrom(e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px"}}/></F>
            <F lbl="To"><input type="date" value={shareTo} onChange={e=>setShareTo(e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px"}}/></F>
          </Grid>
        )}
        <div style={{background:"#f8fafc",borderRadius:"10px",padding:"14px",marginTop:"16px",fontFamily:"monospace",fontSize:"12px",whiteSpace:"pre-wrap",border:"1px solid #e5e7eb",maxHeight:"200px",overflowY:"auto"}}>{getShareData()}</div>
        <div style={{display:"flex",gap:"10px",marginTop:"16px"}}>
          <button onClick={()=>window.open("https://wa.me/?text="+encodeURIComponent(getShareData()),"_blank")} style={{flex:1,padding:"13px",borderRadius:"10px",border:"none",background:"#25d366",color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"700",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px"}}>
            <i className="ti ti-brand-whatsapp" style={{fontSize:"18px"}} aria-hidden/>WhatsApp
          </button>
          <button onClick={()=>window.open("mailto:?subject=SPL DPR Report&body="+encodeURIComponent(getShareData()),"_blank")} style={{flex:1,padding:"13px",borderRadius:"10px",border:"none",background:"#2563eb",color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"700",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px"}}>
            <i className="ti ti-mail" style={{fontSize:"18px"}} aria-hidden/>Email
          </button>
          <button onClick={()=>{navigator.clipboard?.writeText(getShareData()).then(()=>flash("Copied!")).catch(()=>{});}} style={{padding:"13px 16px",borderRadius:"10px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"14px",color:"#374151",fontWeight:"700"}}>Copy</button>
        </div>
        <button onClick={onClose} style={{width:"100%",marginTop:"12px",padding:"12px",borderRadius:"10px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"14px",color:"#6b7280"}}>Close</button>
      </div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App(){
  const [loading,setLoading]=useState(true);
  const [gUser,setGUser]=useState(undefined);   // Google account: undefined=checking, null=signed out, obj=signed in
  const [authReady,setAuthReady]=useState(false);
  const [subs,setSubs]=useState([]);
  const [fbStatus,setFbStatus]=useState('connecting');
  const [engineers,setEngineers]=useState([]);
  const [workTypes,setWorkTypes]=useState(DEFAULT_WT);
  const [users,setUsers]=useState([]);
  const [lateRequests,setLateRequests]=useState([]);
  const [auditTrail,setAuditTrail]=useState([]);
  const [user,setUser]=useState(null);
  const [showLogin,setShowLogin]=useState(false);
  const [appView,setAppView]=useState("welcome"); // welcome | globalAdmin | projects | dashboard
  const [projects,setProjects]=useState([]);
  const [activeProject,setActiveProject]=useState(null);
  const [view,setView]=useState("dashboard");
  const [step,setStep]=useState(0);
  useEffect(()=>{if(view==="form")window.scrollTo({top:0,behavior:"smooth"});},[step,view]);
  const [hdr,setHdr]=useState(mkHdr());
  const [acts,setActs]=useState([]);
  const [matTxs,setMatTxs]=useState([]);
  const [dashDate,setDashDate]=useState(new Date(Date.now()-86400000).toISOString().slice(0,10));
  const [apvDate,setApvDate]=useState(new Date(Date.now()-86400000).toISOString().slice(0,10));
  const [apvId,setApvId]=useState(null);
  const [apvNote,setApvNote]=useState("");
  const [toast,setToast]=useState(null);
  const [newEng,setNewEng]=useState({name:"",dept:"HW - Highway",incharge:INCHARGE_OPTS[0],designation:"Site Engineer"});
  const [newUsr,setNewUsr]=useState({name:"",role:"engineer",pin:"",desc:"",projectAccess:"none",assignedProjectId:"",assignedProjectIds:[],caps:{fill:true,approve:false,download:false,manage:false,settings:false}});
  const [wtEdit,setWtEdit]=useState("");
  const [editActIdx,setEditActIdx]=useState(null);
  const [editMatIdx,setEditMatIdx]=useState(null);
  const [editingSubId,setEditingSubId]=useState(null);
  const [reportFrom,setReportFrom]=useState(()=>{
    // Default to 3 months back to show recent data by default
    const d=new Date();d.setMonth(d.getMonth()-3);
    return d.toISOString().slice(0,10);
  });
  const [reportTo,setReportTo]=useState(()=>new Date().toISOString().slice(0,10));
  const [shareOpen,setShareOpen]=useState(false);
  const [settingsTab,setSettingsTab]=useState("workTypes");
  const [listEdits,setListEdits]=useState({});
  const [copyModal,setCopyModal]=useState(null); // {acts, matTxs, checked}
  const [offlineQueue,setOfflineQueue]=useState(()=>{try{return JSON.parse(localStorage.getItem('dpr_offline_queue')||'[]');}catch(e){return [];}});
  const [isOnline,setIsOnline]=useState(navigator.onLine);
  const [dailyPdfDate,setDailyPdfDate]=useState(null);
  const [lists,setLists]=useState({
    assetGroups:ASSET_GROUPS,
    labourTypes:LABOUR_TYPES,
    vehicleTypes:VEHICLE_TYPES,
    bulkMats:BULK_MATS,
    depts:DEPTS,
    designations:DESIGNATIONS,
    sides:SIDES,
    units:UNITS,
    prodMats:PROD_MATS.map(x=>x.n),
  });
  const [globalLists,setGlobalLists]=useState({
    roles:Object.keys(ROLE_CAPS),
    dateLockDays:2,
    starStart:5,
    lateEngDeduct:0.5,
    lateInchargeDeduct:0.25,
    dprApprovalDays:3,
    requiredFields:{
      // Step 1 - Header
      dept:true, incharge:false, weather:false, difficulty:false,
      // Step 3 - Work Activity
      actDesc:false, chFrom:false, chTo:false, side:false, theoQty:true, prodQty:false, actRemarks:false,
      // Step 2 - Material
      recvCH:false, source:false, recvQty:true, recvTransporter:false, recvLoads:false,
      sendFromCH:false, sendToCH:false, sendQty:true, sendTransporter:false, sendLoads:false, matRemarks:false,
    }
  });
  const [stepErrors,setStepErrors]=useState({});
  const [hasUnsavedForm,setHasUnsavedForm]=useState(false);
  const [installPrompt,setInstallPrompt]=useState(null);
  const [showInstall,setShowInstall]=useState(false); // install banner visible (auto-hides)
  const [iosInstall,setIosInstall]=useState(false); // iOS Safari: show manual add-to-home instructions
  const [showMoreNav,setShowMoreNav]=useState(false); // mobile "More" tabs sheet
  const mobile=useMobile();
  const flash=(m,t="ok")=>{setToast({m,t});setTimeout(()=>setToast(null),3500);};

  // ── PWA install prompt: capture the browser event so we can trigger it from a button ──
  useEffect(()=>{
    const isStandalone=window.matchMedia('(display-mode: standalone)').matches||window.navigator.standalone===true;
    const onPrompt=(e)=>{ e.preventDefault(); if(!isStandalone) setInstallPrompt(e); };
    const onInstalled=()=>{ setInstallPrompt(null); setShowInstall(false); };
    window.addEventListener('beforeinstallprompt',onPrompt);
    window.addEventListener('appinstalled',onInstalled);
    return()=>{ window.removeEventListener('beforeinstallprompt',onPrompt); window.removeEventListener('appinstalled',onInstalled); };
  },[]);

  // After a user signs in (PIN accepted), surface the install prompt briefly at the top.
  useEffect(()=>{
    if(!user)return;
    if(sessionStorage.getItem('splInstallShown')==='1')return;
    const isStandalone=window.matchMedia('(display-mode: standalone)').matches||window.navigator.standalone===true;
    if(isStandalone)return;
    const ua=navigator.userAgent||"";
    const isIOS=/iphone|ipad|ipod/i.test(ua);
    if(installPrompt||isIOS){
      setIosInstall(isIOS&&!installPrompt);
      setShowInstall(true);
      sessionStorage.setItem('splInstallShown','1');
      const t=setTimeout(()=>setShowInstall(false),9000); // auto-hide after a few seconds
      return()=>clearTimeout(t);
    }
  },[user,installPrompt]);

  // ── Android / browser Back → step toward home, then press-again-to-exit ──
  const navRef=useRef({});
  navRef.current={appView,view,activeProject,user};
  const exitArmRef=useRef(0);
  useEffect(()=>{
    if(!user)return;
    try{ history.pushState({spl:1},""); }catch(e){}
    const onPop=()=>{
      const st=navRef.current,u=st.user;
      let atHome=false;
      if(st.activeProject&&st.view!=="dashboard"){ setView("dashboard"); }
      else if(st.activeProject){ // already on the in-project dashboard
        if(u&&(roleKey(u.role)==="admin"||roleKey(u.role)==="management")){ setActiveProject(null); setView("dashboard"); setAppView(roleKey(u.role)==="admin"?"globalAdmin":"projects"); }
        else atHome=true; // engineer / incharge: dashboard is home
      } else {
        const home=(u&&roleKey(u.role)==="admin")?"globalAdmin":"projects";
        if(st.appView!==home){ setAppView(home); }
        else atHome=true;
      }
      if(atHome){
        if(Date.now()-exitArmRef.current<2500){
          window.removeEventListener("popstate",onPop);
          history.back(); // second press within the window → leave the app
          return;
        }
        exitArmRef.current=Date.now();
        flash("Press back again to exit");
      }
      try{ history.pushState({spl:1},""); }catch(e){} // re-arm the trap so Back stays captured
    };
    window.addEventListener("popstate",onPop);
    return ()=>window.removeEventListener("popstate",onPop);
  },[user]);

  async function doInstall(){
    if(!installPrompt)return;
    setShowInstall(false);
    installPrompt.prompt();
    try{ await installPrompt.userChoice; }catch(e){}
    setInstallPrompt(null);
  }

  // Install banner — slides in at the very top after login, auto-hides after a few seconds.
  const installBanner=(showInstall&&(installPrompt||iosInstall))?(
    <div style={{position:"fixed",top:0,left:0,right:0,zIndex:100000,display:"flex",justifyContent:"center",padding:"10px",pointerEvents:"none"}}>
      <div style={{pointerEvents:"auto",display:"flex",alignItems:"center",gap:"12px",background:"#fff",border:`1px solid #e5e7eb`,borderLeft:`4px solid ${AM}`,borderRadius:"12px",boxShadow:"0 8px 30px rgba(0,0,0,.18)",padding:"10px 12px 10px 14px",maxWidth:"440px",width:"100%",animation:"none"}}>
        <img src={LOGO} alt="SPL DPR" style={{height:"34px",width:"34px",objectFit:"contain",borderRadius:"8px",background:"#fff",border:"1px solid #eef2f7",flexShrink:0}}/>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:"13px",fontWeight:"800",color:NV,lineHeight:1.2}}>Install SPL DPR</div>
          <div style={{fontSize:"11px",color:"#6b7280",fontWeight:"600"}}>{iosInstall?<span>Tap <i className="ti ti-share-2" style={{fontSize:"12px",verticalAlign:"-1px"}} aria-hidden/> Share, then “Add to Home Screen”</span>:"Add to your home screen for quick access"}</div>
        </div>
        {!iosInstall&&<button onClick={doInstall} style={{flexShrink:0,padding:"9px 15px",borderRadius:"9px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"800",display:"flex",alignItems:"center",gap:"6px"}}>
          <i className="ti ti-download" style={{fontSize:"15px"}} aria-hidden/>Install
        </button>}
        <button onClick={()=>setShowInstall(false)} aria-label="Dismiss" style={{flexShrink:0,width:"32px",height:"32px",borderRadius:"8px",border:"none",background:"#f1f5f9",color:"#64748b",cursor:"pointer",fontSize:"15px"}}>
          <i className="ti ti-x" aria-hidden/>
        </button>
      </div>
    </div>
  ):null;

  // ── Google Auth gate: watch sign-in state ────────────────────────────────
  useEffect(()=>{
    const unsub=onAuthStateChanged(auth,u=>{ setGUser(u||null); setAuthReady(true); });
    return()=>unsub();
  },[]);
  async function handleGoogleSignIn(){
    try{ await signInWithPopup(auth,googleProvider); }
    catch(e){ flash("Google sign-in failed: "+(e.code||e.message),"err"); }
  }
  async function handleGoogleSignOut(){
    setUser(null);setActiveProject(null);setView("dashboard");setAppView("welcome");
    try{ await fbSignOut(auth); }catch(e){}
  }

  // Warn before page refresh when DPR form has data
  useEffect(()=>{
    const h=(e)=>{
      if(view==="form"&&(acts.length>0||matTxs.length>0||hdr.engineer)){
        e.preventDefault();e.returnValue="You have an unsaved DPR. Leave anyway?";return e.returnValue;
      }
    };
    window.addEventListener('beforeunload',h);
    return()=>window.removeEventListener('beforeunload',h);
  },[view,acts,matTxs,hdr.engineer]);

  // Auto-fill incharge and dept when engineers data loads (handles race condition)
  useEffect(()=>{
    if(view==="form"&&roleKey(user?.role)==="engineer"&&engineers.length>0){
      if(!hdr.incharge||!hdr.dept){
        const engRec=engineers.find(x=>
          x.id===user.id||
          x.name===user.name||
          (x.name||"").trim().toLowerCase()===(user.name||"").trim().toLowerCase()
        );
        if(engRec){
          setHdr(p=>({
            ...p,
            incharge:p.incharge||(engRec.incharge||""),
            dept:p.dept||(engRec.dept||""),
          }));
        }
      }
    }
  },[engineers,view,user]);

  // Auto-save DPR draft to localStorage as engineer fills
  useEffect(()=>{
    if(view==="form"){
      localStorage.setItem('dpr_draft',JSON.stringify({hdr,acts,matTxs,step,savedAt:new Date().toISOString()}));
    }
  },[hdr,acts,matTxs,step,view]);

  // ── GLOBAL effect: users + projects list (runs once) ────────────────────
  // ── Online/Offline detection + queue sync ───────────────────────────────
  useEffect(()=>{
    const handleOnline=()=>{
      setIsOnline(true);
      const queue=JSON.parse(localStorage.getItem('dpr_offline_queue')||'[]');
      if(queue.length>0){
        flash("📡 Back online — syncing "+queue.length+" queued DPR(s)...");
        Promise.all(queue.map(async sub=>
          fetch(await authedUrl(RTDB_URL+'/projects/'+sub._projectId+'/submissions/'+sub.id+'.json'),{
            method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(sub)
          })
        )).then(()=>{
          localStorage.removeItem('dpr_offline_queue');
          setOfflineQueue([]);
          flash("✅ All queued DPRs synced!");
        }).catch(()=>flash("⚠️ Sync failed — will retry","err"));
      }
    };
    const handleOffline=()=>{setIsOnline(false);flash("📵 No internet — DPRs will be saved locally and synced when you're back online","err");};
    window.addEventListener('online',handleOnline);
    window.addEventListener('offline',handleOffline);
    return()=>{window.removeEventListener('online',handleOnline);window.removeEventListener('offline',handleOffline);};
  },[]);

  useEffect(()=>{
    setLoading(false);
    // ⚠️ Gate on Google auth: DB rules require auth != null. Attaching onValue
    // listeners before sign-in completes gets permission_denied, which the SDK
    // treats as a permanent cancel — the listener never re-fires after login,
    // so freshly-created projects (written via authenticated REST) never stream
    // back. Attach only once we have a signed-in Google user, and re-run if it
    // changes, so listeners always carry a valid auth token.
    if(!gUser) return;
    const _unsubs=[];
    // Users
    const usersRef=ref(db,'users');
    const u3=onValue(usersRef,snap=>{
      const obj=snap.val()||{};
      const arr=Object.values(obj);
      setUsers(arr.length>0?arr:[]);
    },err=>console.error('Users:',err.code));
    _unsubs.push(()=>{try{u3();}catch(e){}});
    // Projects list
    const projRef=ref(db,'projects');
    const up=onValue(projRef,snap=>{
      const obj=snap.val()||{};
      const arr=Object.values(obj).filter(p=>p&&p.id);
      arr.sort((a,b)=>(a.name||'').localeCompare(b.name||''));
      setProjects(arr);
    },err=>console.error('Projects:',err.code));
    _unsubs.push(()=>{try{up();}catch(e){}});
    // Global lists (roles, designations, depts, incharge names)
    const glRef=ref(db,'config/globalLists');
    const ugl=onValue(glRef,snap=>{
      const dt=snap.val()||{};
      setGlobalLists(prev=>({
        ...prev,
        ...(dt.roles&&{roles:dt.roles}),
        ...(dt.dateLockDays!=null&&{dateLockDays:Number(dt.dateLockDays)}),
        ...(dt.starStart!=null&&{starStart:Number(dt.starStart)}),
        ...(dt.lateEngDeduct!=null&&{lateEngDeduct:Number(dt.lateEngDeduct)}),
        ...(dt.lateInchargeDeduct!=null&&{lateInchargeDeduct:Number(dt.lateInchargeDeduct)}),
        ...(dt.dprApprovalDays!=null&&{dprApprovalDays:Number(dt.dprApprovalDays)}),
        ...(dt.requiredFields&&{requiredFields:{...prev.requiredFields,...dt.requiredFields}}),
      }));
    });
    _unsubs.push(()=>{try{ugl();}catch(e){}});
    // Late-entry requests (backlog approval workflow)
    const lrRef=ref(db,'lateRequests');
    const ulr=onValue(lrRef,snap=>{
      const obj=snap.val()||{};
      const arr=Object.values(obj).filter(Boolean);
      arr.sort((a,b)=>(b.requestedAt||'').localeCompare(a.requestedAt||''));
      setLateRequests(arr);
    },err=>console.error('LateRequests:',err.code));
    _unsubs.push(()=>{try{ulr();}catch(e){}});
    // Audit trail (late-submission & backlog-approval star penalties)
    const alRef=ref(db,'auditLog');
    const ual=onValue(alRef,snap=>{
      const obj=snap.val()||{};
      const arr=Object.values(obj).filter(Boolean);
      arr.sort((a,b)=>(b.ts||'').localeCompare(a.ts||''));
      setAuditTrail(arr);
    },err=>console.error('AuditLog:',err.code));
    _unsubs.push(()=>{try{ual();}catch(e){}});
    return()=>_unsubs.forEach(fn=>fn());
  },[gUser]);

  // ── PROJECT-SCOPED effect: runs when activeProject changes ───────────────
  useEffect(()=>{
    if(!activeProject){setSubs([]);setEngineers([]);return;}
    const pid=activeProject.id;
    const base='projects/'+pid;
    const _unsubs=[];

    // Submissions — REST poll + SDK realtime
    let sdkConnected=false;
    async function fetchSubs(){
      if(sdkConnected)return; // SDK is live, no need to poll
      try{
        const r=await fetch(await authedUrl(RTDB_URL+'/'+base+'/submissions.json'));
        if(!r.ok) throw new Error('HTTP '+r.status);
        const obj=await r.json();
        if(obj){
          const arr=Object.values(obj).map(normalizeSub);
          arr.sort((a,b)=>(b.submittedAt||'').localeCompare(a.submittedAt||''));
          setSubs(arr); setFbStatus('ok');
        } else { setSubs([]); setFbStatus('ok'); }
      }catch(e){ console.error('Subs fetch:',e.message); setFbStatus('error'); }
    }
    fetchSubs();
    const pollTimer=setInterval(fetchSubs, 8000);
    _unsubs.push(()=>clearInterval(pollTimer));
    const subsRef=ref(db,base+'/submissions');
    const u1=onValue(subsRef,snap=>{
      sdkConnected=true; // SDK connected — stop polling
      const obj=snap.val()||{};
      const arr=Object.values(obj).map(normalizeSub);
      arr.sort((a,b)=>(b.submittedAt||'').localeCompare(a.submittedAt||''));
      setSubs(arr); setFbStatus('ok');
    },err=>{ sdkConnected=false; console.error('Subs SDK:',err.code); });
    _unsubs.push(()=>{try{u1();}catch(e){}});

    // Engineers
    async function fetchEngs(){
      try{
        const r=await fetch(await authedUrl(RTDB_URL+'/'+base+'/engineers.json'));
        const obj=await r.json();
        if(obj){const arr=Object.values(obj);setEngineers(arr);}
        else setEngineers([]);
      }catch(e){console.error('Engs:',e.message);}
    }
    fetchEngs();
    const engsRef=ref(db,base+'/engineers');
    const u2=onValue(engsRef,snap=>{
      const obj=snap.val()||{};
      const arr=Object.values(obj);
      setEngineers(arr.length>0?arr:[]);
    },err=>console.error('Engs SDK:',err.code));
    _unsubs.push(()=>{try{u2();}catch(e){}});

    // Work types
    const wtRef=ref(db,base+'/config/workTypes');
    const u4=onValue(wtRef,snap=>{
      const v=snap.val();
      if(v&&v.types) setWorkTypes(v.types);
      else setWorkTypes(DEFAULT_WT);
    });
    _unsubs.push(()=>{try{u4();}catch(e){}});

    // Editable lists
    const listsRef=ref(db,base+'/config/editableLists');
    const u5=onValue(listsRef,snap=>{
      const dt=snap.val()||{};
      if(Object.keys(dt).length>0){
        setLists(prev=>({...prev,
          assetGroups:dt.assetGroups||prev.assetGroups,
          labourTypes:dt.labourTypes||prev.labourTypes,
          vehicleTypes:dt.vehicleTypes||prev.vehicleTypes,
          bulkMats:dt.bulkMats||prev.bulkMats,
          depts:dt.depts||prev.depts,
          designations:dt.designations||prev.designations,
          sides:dt.sides||prev.sides,
          units:dt.units||prev.units,
          prodMats:dt.prodMats||prev.prodMats,
        }));
      }
    });
    _unsubs.push(()=>{try{u5();}catch(e){}});

    return()=>_unsubs.forEach(fn=>fn());
  },[activeProject]);

  const caps=user?(user.caps||ROLE_CAPS[roleKey(user.role)]||ROLE_CAPS.engineer):ROLE_CAPS.engineer;
  if(loading)return <div style={{padding:"5rem",textAlign:"center",color:"#6b7280",fontFamily:"sans-serif",fontSize:"16px"}}>Loading…</div>;

  // ── GOOGLE SIGN-IN GATE ──────────────────────────────────────────────
  // The database is locked to authenticated Google users. Everyone must sign in
  // with Google FIRST; the per-account PIN gate follows on the welcome screen.
  if(!authReady)return <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:NV,color:"#fff",fontFamily:"var(--font-sans)",fontSize:"15px"}}>Checking sign-in…</div>;
  if(!gUser)return(
    <div style={{fontFamily:"var(--font-sans)",minHeight:"100vh",background:`linear-gradient(145deg,${NV} 0%,#2d5a9e 50%,#1a4080 100%)`,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"24px",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:"-80px",right:"-80px",width:"320px",height:"320px",borderRadius:"50%",background:"rgba(255,255,255,.05)"}}/>
      <div style={{position:"absolute",bottom:"-60px",left:"-60px",width:"260px",height:"260px",borderRadius:"50%",background:"rgba(255,255,255,.04)"}}/>
      {toast&&<div style={{position:"fixed",top:"14px",left:"50%",transform:"translateX(-50%)",background:toast.t==="err"?"#fef2f2":"#f0fdf4",color:toast.t==="err"?RD:GN,padding:"12px 22px",borderRadius:"10px",boxShadow:"0 4px 20px rgba(0,0,0,.18)",zIndex:99998,fontSize:"14px",fontWeight:"700",border:`2px solid ${toast.t==="err"?"#fca5a5":"#86efac"}`,display:"flex",alignItems:"center",gap:"8px",maxWidth:"90vw"}}><i className={`ti ti-${toast.t==="err"?"alert-circle":"circle-check"}`} aria-hidden/>{toast.m}</div>}
      <div style={{textAlign:"center",maxWidth:"420px",zIndex:1}}>
        <img src={LOGO} alt="SPL" style={{height:"70px",borderRadius:"10px",background:"#fff",padding:"6px",marginBottom:"24px",boxShadow:"0 8px 30px rgba(0,0,0,.3)"}}/>
        <div style={{fontSize:"24px",fontWeight:"800",color:"#fff",marginBottom:"6px",letterSpacing:"-0.5px"}}>SPL Infrastructure Pvt. Ltd.</div>
        <div style={{fontSize:"14px",color:"rgba(255,255,255,.6)",marginBottom:"36px",letterSpacing:"0.05em",textTransform:"uppercase"}}>Daily Progress Report System</div>
        <div style={{background:"rgba(255,255,255,.08)",borderRadius:"14px",padding:"22px 24px",marginBottom:"28px",border:"1px solid rgba(255,255,255,.12)"}}>
          <i className="ti ti-shield-lock" style={{fontSize:"30px",color:"rgba(255,255,255,.8)"}} aria-hidden/>
          <div style={{fontSize:"14px",color:"rgba(255,255,255,.85)",lineHeight:"1.6",marginTop:"10px"}}>This is a private, authorised-access app. Sign in with your Google account to continue — you'll then enter your personal PIN.</div>
        </div>
        <button onClick={handleGoogleSignIn} style={{width:"100%",maxWidth:"340px",padding:"15px 18px",borderRadius:"14px",border:"none",background:"#fff",color:"#1f2937",cursor:"pointer",fontSize:"16px",fontWeight:"700",boxShadow:"0 4px 20px rgba(0,0,0,.25)",display:"flex",alignItems:"center",justifyContent:"center",gap:"12px"}}>
          <svg width="20" height="20" viewBox="0 0 48 48" aria-hidden><path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/><path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/><path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/><path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/></svg>
          Continue with Google
        </button>
        <div style={{fontSize:"12px",color:"rgba(255,255,255,.35)",marginTop:"18px"}}>Access is restricted to authorised SPL accounts</div>
      </div>
    </div>
  );

  const PAD=mobile?"12px":"16px";
  const CONTENT_PAD=mobile?"80px":"16px";
  const C={background:"#fff",border:"1px solid #e5e7eb",borderRadius:"14px",padding:"16px"};
  const SH={fontSize:"12px",fontWeight:"700",color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:"14px",paddingBottom:"10px",borderBottom:"1px solid #f3f4f6",display:"flex",alignItems:"center",justifyContent:"space-between"};
  // saveSubs → replaced by direct Firestore calls below
  
  const saveWt=async w=>{
    if(!activeProject)return;
    setWorkTypes(w);
    rtdbPut('projects/'+activeProject.id+'/config/workTypes',{types:w})
      .then(()=>flash('Saved — '+w.length+' items'))
      .catch(e=>flash('Failed: '+e.message,'err'));
  };

  function loadSubForEdit(sub){
    setHdr({engineer:sub.engineer,date:sub.date,shift:sub.shift,dept:sub.dept,incharge:sub.incharge,engCustom:"",weather:sub.weather||"☀️ Clear",difficulty:sub.difficulty||""});
    setActs(sub.activities||[]);
    setMatTxs(sub.matTxs||[]);
    setEditingSubId(sub.id);
    setStep(0);
    setView("form");
    flash("📝 Editing your DPR — make corrections then re-submit");
  }

  function updHdr(f,v){setHdr(p=>{const n={...p,[f]:v};if(f==="engineer"&&v){const e=engineers.find(x=>x.name===v||(x.name||"").trim().toLowerCase()===v.trim().toLowerCase());if(e){if(e.incharge)n.incharge=e.incharge;if(e.dept)n.dept=e.dept;}}return n;});}

  // When entering form view, pre-fill name if logged in as engineer
  // Guard navigation away from form when unsaved data exists
  function safeNav(fn){
    if(view==="form"&&(acts.length>0||matTxs.length>0||hasUnsavedForm)){
      if(!window.confirm("⚠️ You have unsaved DPR data (activities or materials).\n\nLeave and lose all this data?"))return;
    }
    setHasUnsavedForm(false);
    fn();
  }

  function openForm(){
    try{
      const saved=JSON.parse(localStorage.getItem('dpr_draft')||'null');
      if(saved&&saved.hdr){
        const minsAgo=(Date.now()-new Date(saved.savedAt).getTime())/60000;
        if(minsAgo<120&&(saved.acts?.length>0||saved.matTxs?.length>0)){
          if(window.confirm("📋 You have an unsaved DPR draft from "+Math.round(minsAgo)+" min ago.\n\nRestore it?")){
            setHdr(saved.hdr);setActs(saved.acts||[]);setMatTxs(saved.matTxs||[]);setStep(saved.step||0);
            setView("form");return;
          }
        }
      }
    }catch(e){}
    // Fresh form — pre-fill from engineer record
    if(roleKey(user?.role)==="engineer"){
      // Match by user.id (primary key), name, or name trimmed+lowercase
      const engRec=engineers.find(x=>
        x.id===user.id ||
        x.name===user.name ||
        (x.name||"").trim().toLowerCase()===(user.name||"").trim().toLowerCase()
      );
      const incharge=engRec?.incharge||"";
      const dept=engRec?.dept||"";
      setHdr(()=>({
        ...mkHdr(),
        engineer:user.name,
        incharge,
        dept,
      }));
      if(!incharge){
        flash("⚠️ No Incharge assigned. Ask admin to set it in the Engineers tab.","err");
      }
    } else {
      setHdr(mkHdr());
    }
    setStep(0);
    setActs([]);
    setMatTxs([]);
    setHasUnsavedForm(false);
    setView("form");
  }

  function submitDPR(){
    const eng=roleKey(user?.role)==="engineer"?user.name:(hdr.engineer==="__other"?(hdr.engCustom||"").trim():hdr.engineer);
    if(!eng){ flash("Please select your name","err"); return; }
    if(acts.length===0&&matTxs.length===0){ flash("Add at least one work activity or material entry","err"); return; }
    if(!hdr.incharge){if(!window.confirm("⚠️ No Incharge assigned.\n\nProceed anyway?"))return;}
    // Date lock check (admin-configurable) — engineers/incharges need an approved
    // late-entry request to submit past the lock window; admin is never restricted.
    const maxDays=globalLists.dateLockDays!=null?Number(globalLists.dateLockDays):2;
    let viaApprovedRequest=null;
    if(maxDays>0&&roleKey(user?.role)!=="admin"){
      const daysAgo=Math.floor((new Date()-new Date(hdr.date+"T12:00:00"))/86400000);
      if(daysAgo>maxDays){
        viaApprovedRequest=lateRequests.find(r=>r.status==="approved"&&r.projectId===activeProject.id&&r.engName===eng&&r.date===hdr.date);
        if(!viaApprovedRequest){
          const pendingReq=lateRequests.find(r=>r.status==="pending"&&r.projectId===activeProject.id&&r.engName===eng&&r.date===hdr.date);
          if(pendingReq){flash("⏳ Your late-entry request for "+hdr.date+" is still awaiting admin approval.","err");return;}
          if(window.confirm("❌ Date locked: DPRs older than "+maxDays+" day"+(maxDays===1?"":"s")+" need admin approval.\n\nSend a late-entry request for "+hdr.date+" now?")){
            const reason=window.prompt("Briefly explain why this entry is late:","");
            if(reason!=null)raiseLateRequest(eng,hdr.date,reason||"(no reason given)");
          }
          return;
        }
      }
    }
    if(!editingSubId){
      const dup=subs.find(s=>s.engineer===eng&&s.date===hdr.date&&s.shift===hdr.shift);
      if(dup&&!window.confirm("⚠️ You already have a "+hdr.shift+" Shift DPR for "+hdr.date+".\n\nSubmit as an additional entry?"))return;
    }
    const isRevisionEdit=!!editingSubId;
    // Build edit history entry
    const existingSub=editingSubId?subs.find(s=>s.id===editingSubId):null;
    const historyEntry=isRevisionEdit?{editedAt:new Date().toISOString(),editedBy:user?.name||eng,prevActCount:(existingSub?.activities||[]).length,newActCount:acts.length,prevMatCount:(existingSub?.matTxs||[]).length,newMatCount:matTxs.length}:null;
    const sub={
      id:editingSubId||uid(), submittedAt:new Date().toISOString(),
      engineer:eng, date:hdr.date, shift:hdr.shift,
      dept:hdr.dept, incharge:hdr.incharge,
      weather:hdr.weather||"☀️ Clear", difficulty:hdr.difficulty||"",
      approved:false, approvedBy:"", approvedAt:"", approvalNote:"",
      needsRevision:false, revisionNote:"",
      activities:[...acts], matTxs:[...matTxs],
      _projectId:activeProject?.id,
      editHistory:historyEntry?[...(existingSub?.editHistory||[]),historyEntry]:[],
    };
    const savedDate=hdr.date;
    setHdr(mkHdr()); setActs([]); setMatTxs([]); setStep(0); setEditingSubId(null);
    setView("dashboard"); setDashDate(savedDate);
    localStorage.removeItem('dpr_draft'); // clear auto-save on submit
    if(!isOnline){
      const q=[...offlineQueue,sub];
      localStorage.setItem('dpr_offline_queue',JSON.stringify(q));
      setOfflineQueue(q);
      flash("📵 Saved offline — will sync when connected ("+q.length+" queued)","ok");
      return;
    }
    flash(isRevisionEdit?"⏳ Saving revised DPR…":"⏳ Saving DPR...");
    authedUrl(RTDB_URL+'/projects/'+activeProject.id+'/submissions/'+sub.id+'.json').then(_u=>fetch(_u,{
      method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(sub)
    })).then(r=>{
      if(!r.ok)return r.text().then(t=>{throw new Error(t);});
      setFbStatus('ok');
      flash(isRevisionEdit?"✅ Revised DPR saved!":"✅ DPR saved!");
      // Performance: a fresh (non-revision) submission filed on a day other than
      // its own date is "late" — deduct stars from the engineer + their incharge
      // and record it in the Audit Log. If it only made it in because an admin
      // approved a late-entry request, log it as a backlog approval instead.
      if(!isRevisionEdit&&savedDate!==new Date().toISOString().slice(0,10)){
        const engDeduct=Number(globalLists.lateEngDeduct??0.5);
        const icDeduct=Number(globalLists.lateInchargeDeduct??0.25);
        adjustStars(eng,-engDeduct);
        if(hdr.incharge)adjustStars(hdr.incharge,-icDeduct);
        if(viaApprovedRequest){
          logAudit({type:'backlog_approved',engName:eng,incharge:hdr.incharge||'',approver:viaApprovedRequest.decidedBy||'Admin',engDeduct,icDeduct:hdr.incharge?icDeduct:0,detail:'Backlog entry for '+savedDate+' — '+(viaApprovedRequest.reason||'no reason given'),projectId:activeProject.id,projectName:activeProject.name,date:savedDate});
        }else{
          logAudit({type:'late_submit',engName:eng,incharge:hdr.incharge||'',approver:'System',engDeduct,icDeduct:hdr.incharge?icDeduct:0,detail:'DPR for '+savedDate+' filed late',projectId:activeProject.id,projectName:activeProject.name,date:savedDate});
        }
      }
      if(!isRevisionEdit&&hdr.incharge){
        const icUser=users.find(u=>u.name===hdr.incharge&&u.phone);
        if(icUser?.phone){
          const msg=`📋 *New DPR Submitted*\n👷 ${eng}\n📅 ${hdr.date} (${hdr.shift})\n🔖 ${activeProject.code||activeProject.name}\n🔧 ${acts.length} activit${acts.length===1?"y":"ies"}\n\nPlease review and approve.`;
          const phone=icUser.phone.replace(/\D/g,"");
          if(window.confirm("Notify "+hdr.incharge+" via WhatsApp?"))window.open("https://wa.me/"+phone+"?text="+encodeURIComponent(msg),"_blank");
        }
      }
    }).catch(e=>{
      const q=[...offlineQueue,sub];
      localStorage.setItem('dpr_offline_queue',JSON.stringify(q));
      setOfflineQueue(q);
      flash("⚠️ Network error — DPR queued offline","err");
    });
  }

  function canApprove(s){if(!user)return false;if(roleKey(user.role)==="admin"||roleKey(user.role)==="management")return true;if(roleKey(user.role)==="incharge"){const e=engineers.find(x=>x.name===s.engineer);return e&&e.incharge===user.name;}return false;}

  // Reassign user to a new project — single source of truth: user.assignedProjectId
  function reassignUser(u,newPid){
    rtdbPatch('users/'+u.id,{assignedProjectId:newPid||""})
      .then(()=>flash("✅ Project assigned"))
      .catch(e=>flash(e.message,"err"));
  }

  // Role-based routing after login
  function handleLogin(u){
    setUser(u);
    setShowLogin(false);
    flash("✅ Signed in as "+u.name);
    const role=u.role;
    if(role==="admin"){setAppView("globalAdmin");return;}
    if(role==="management"){setAppView("projects");return;}
    // engineer or incharge — route to their single assigned project
    const assignedPid=u.assignedProjectId;
    if(!assignedPid){setAppView("noProject");return;}
    const proj=projects.find(p=>p.id===assignedPid);
    if(!proj){setAppView("noProject");return;}
    setActiveProject(proj);
    setView("dashboard");
    setAppView("dashboard");
  }

  function handleSignOut(){
    setUser(null);
    setActiveProject(null);
    setView("dashboard");
    setAppView("welcome");
  }

  // helper: project base path
  const pb=()=>activeProject?'projects/'+activeProject.id:null;

  async function approveSub(id){if(!user||!pb())return;setApvId(null);setApvNote("");
    // DPR approval window — approving past the admin-set window is a backlog approval
    const subA=subs.find(s=>s.id===id);
    const winDays=Number(globalLists.dprApprovalDays??0);
    if(subA&&winDays>0&&subA.date){
      const age=Math.floor((Date.now()-new Date(subA.date+"T12:00:00").getTime())/86400000);
      if(age>winDays){
        const icDeduct=Number(globalLists.lateInchargeDeduct??0.25);
        if(roleKey(user.role)==="incharge")adjustStars(user.name,-icDeduct);
        logAudit({type:'backlog_approved',engName:subA.engineer||'',incharge:subA.incharge||'',approver:user.name,engDeduct:0,icDeduct:roleKey(user.role)==="incharge"?icDeduct:0,detail:'DPR dated '+subA.date+' approved '+age+' days later — outside the '+winDays+'-day approval window',projectId:activeProject.id,projectName:activeProject.name});
      }
    }
    flash("✅ Approved by "+user.name);rtdbPatch(pb()+'/submissions/'+id,{approved:true,approvedBy:user.name,approvedRole:user.role,approvedAt:new Date().toISOString(),approvalNote:apvNote}).catch(e=>flash('Approve write failed: '+e.message,'err'));}
  async function deleteSub(id){if(!user||roleKey(user.role)!=='admin'||!pb())return;if(!window.confirm('Delete this DPR permanently? This cannot be undone.'))return;rtdbDelete(pb()+'/submissions/'+id).then(()=>flash('🗑 DPR deleted')).catch(e=>flash('Delete failed: '+e.message,'err'));}
  async function addEngineer(){if(!newEng.name.trim()){flash('Enter name','err');return;}if(!pb())return;const nId=uid();rtdbPut(pb()+'/engineers/'+nId,{...newEng,id:nId}).then(()=>flash('Engineer added')).catch(e=>flash('Failed: '+e.message,'err'));setNewEng({name:'',dept:'HW - Highway',incharge:lists.incharge?.[0]||INCHARGE_OPTS[0],designation:"Site Engineer"});}
  async function saveEngEdit(id,u){if(!pb())return;rtdbPatch(pb()+'/engineers/'+id,u).then(()=>flash('Saved')).catch(e=>flash('Failed: '+e.message,'err'));}
  async function addUser(){
    if(!newUsr.name.trim()||!newUsr.pin.trim()){flash("Enter name & PIN","err");return;}
    if(newUsr.pin.length<4){flash("PIN must be 4+ digits","err");return;}
    if(users.find(u=>u.pin===newUsr.pin)){flash("PIN already in use","err");return;}
    const nId2=uid();
    const access=newUsr.projectAccess||"none";
    const newUser={...newUsr,id:nId2,desc:(newUsr.desc||"").trim(),projectAccess:access,caps:newUsr.caps||ROLE_CAPS[roleKey(newUsr.role)]||ROLE_CAPS.engineer,createdAt:new Date().toISOString()};
    if(access!=="specific"){newUser.assignedProjectId="";newUser.assignedProjectIds=[];}
    await rtdbPut('users/'+nId2,newUser).catch(e=>flash('Failed: '+e.message,'err'));
    setNewUsr({name:"",role:"engineer",pin:"",desc:"",projectAccess:"none",assignedProjectId:"",assignedProjectIds:[],caps:{fill:true,approve:false,download:false,manage:false,settings:false}});
    flash("✅ User created");
  }
  async function saveUserCaps(uid_,caps){rtdbPatch('users/'+uid_,{caps}).then(()=>flash("Permissions updated")).catch(e=>flash("Failed: "+e.message,"err"));}

  // ── Performance (stars) + Audit Log + Late-Entry Requests ──────────────────
  function logAudit(entry){
    const id=uid();
    rtdbPut('auditLog/'+id,{id,ts:new Date().toISOString(),...entry}).catch(e=>console.error('Audit log failed:',e.message));
  }
  function adjustStars(name,delta){
    if(!name||!delta)return;
    const u=users.find(x=>x.name===name);
    if(!u)return;
    const cur=u.stars!=null?Number(u.stars):Number(globalLists.starStart??5);
    const next=Math.max(0,Math.round((cur+delta)*100)/100);
    rtdbPatch('users/'+u.id,{stars:next}).catch(e=>console.error('Star update failed:',e.message));
  }
  function raiseLateRequest(engName,date,reason){
    if(!activeProject)return;
    const id=uid();
    rtdbPut('lateRequests/'+id,{
      id,projectId:activeProject.id,projectName:activeProject.name,
      engName,date,reason,status:'pending',requestedAt:new Date().toISOString(),
    }).then(()=>flash("📨 Late-entry request sent to admin for "+date)).catch(e=>flash("Failed: "+e.message,"err"));
  }
  function decideLateRequest(reqId,approve){
    const r=lateRequests.find(x=>x.id===reqId);
    if(!r)return;
    rtdbPatch('lateRequests/'+reqId,{status:approve?'approved':'rejected',decidedBy:user?.name||'Admin',decidedAt:new Date().toISOString()})
      .then(()=>flash(approve?"✅ Approved — "+r.engName+" can now submit "+r.date:"Request rejected"))
      .catch(e=>flash("Failed: "+e.message,"err"));
    // Star penalty + audit entry are applied when the DPR is actually submitted
    // (submitDPR), where the real incharge for that submission is known.
  }
  async function rejectSub(id,note){
    const n=note||apvNote;
    setApvId(null);setApvNote('');
    flash('Marked for Revision.');
    fetch(await authedUrl(RTDB_URL+'/projects/'+activeProject.id+'/submissions/'+id+'.json'),{
      method:'PATCH',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({approved:false,approvedBy:'',approvedAt:'',needsRevision:true,revisionNote:n})
    }).catch(e=>flash('Reject write failed: '+e.message,'err'));
  }
  // (functions now defined above with project-scoped paths)

  const dayS=subs.filter(s=>s.date===dashDate);
  const presentN=new Set(dayS.map(s=>s.engineer));
  const pendN=(()=>{
    const allPending=subs.filter(s=>!s.approved&&!s.needsRevision);
    if(!user)return 0;
    if(roleKey(user.role)==="admin"||roleKey(user.role)==="management")return allPending.length;
    if(roleKey(user.role)==="incharge"){
      const myEngs=engineers.filter(e=>e.incharge===user.name).map(e=>e.name);
      return allPending.filter(s=>myEngs.includes(s.engineer)).length;
    }
    return 0;
  })();

  // Role-scoped filter for Reports tab
  function getReportSubs(){
    const base=subs.filter(s=>s.date>=reportFrom&&s.date<=reportTo);
    if(!user)return base;
    if(roleKey(user.role)==="engineer") return base.filter(s=>s.engineer===user.name);
    if(roleKey(user.role)==="incharge"){
      const myEngs=engineers.filter(e=>e.incharge===user.name).map(e=>e.name);
      return base.filter(s=>myEngs.includes(s.engineer));
    }
    return base; // management, admin — see all
  }
  const flatA=dayS.flatMap(s=>(s.activities||[]).map(a=>({...a,_eng:s.engineer,_dept:s.dept,_shift:s.shift,_ok:s.approved,_by:s.approvedBy})));
  const flatM=dayS.flatMap(s=>(s.matTxs||[]).map(m=>({...m,_eng:s.engineer})));
  const apvS=subs.filter(s=>s.date===apvDate);

  // Revision submissions this engineer needs to fix
  const myRevisions=roleKey(user?.role)==="engineer"?subs.filter(s=>s.engineer===user.name&&s.needsRevision&&!s.approved):[];
  // This engineer's own late-entry requests (so they know their status)
  const myLateRequests=roleKey(user?.role)==="engineer"?lateRequests.filter(r=>r.engName===user.name&&r.projectId===activeProject?.id).sort((a,b)=>(b.requestedAt||"").localeCompare(a.requestedAt||"")):[];

  // TABS — filter by role caps
  const allTabs=[
    {id:"dashboard",i:"ti-layout-dashboard",l:"Dashboard"},
    {id:"form",i:"ti-pencil-plus",l:"Fill DPR"},
    {id:"approve",i:"ti-circle-check",l:"Approve",badge:pendN,hide:!caps.approve},
    {id:"monthly",i:"ti-calendar-stats",l:"Reports"},
    {id:"manage",i:"ti-users-group",l:"Engineers",hide:!caps.manage},
    {id:"settings",i:"ti-settings",l:"Settings",hide:!caps.settings},
  ].filter(t=>!t.hide);

  // ── WELCOME SCREEN ───────────────────────────────────────────────────────
  if(appView==="welcome"){
    const q=QUOTES[new Date().getDate()%QUOTES.length];
    return(
      <div style={{fontFamily:"var(--font-sans)",minHeight:"100vh",background:`linear-gradient(145deg,${NV} 0%,#2d5a9e 50%,#1a4080 100%)`,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"24px",position:"relative",overflow:"hidden"}}>
        {/* Background decoration */}
        <div style={{position:"absolute",top:"-80px",right:"-80px",width:"320px",height:"320px",borderRadius:"50%",background:"rgba(255,255,255,.05)"}}/>
        <div style={{position:"absolute",bottom:"-60px",left:"-60px",width:"260px",height:"260px",borderRadius:"50%",background:"rgba(255,255,255,.04)"}}/>
        {toast&&<div style={{position:"fixed",top:"14px",left:"50%",transform:"translateX(-50%)",background:toast.t==="err"?"#fef2f2":"#f0fdf4",color:toast.t==="err"?RD:GN,padding:"12px 22px",borderRadius:"10px",boxShadow:"0 4px 20px rgba(0,0,0,.18)",zIndex:99998,fontSize:"14px",fontWeight:"700",border:`2px solid ${toast.t==="err"?"#fca5a5":"#86efac"}`,display:"flex",alignItems:"center",gap:"8px",maxWidth:"90vw"}}><i className={`ti ti-${toast.t==="err"?"alert-circle":"circle-check"}`} aria-hidden/>{toast.m}</div>}
        {showLogin&&<LoginModal users={users} onLogin={handleLogin} onClose={()=>setShowLogin(false)}/>}
        <div style={{textAlign:"center",maxWidth:"480px",zIndex:1}}>
          <img src={LOGO} alt="SPL" style={{height:"70px",borderRadius:"10px",background:"#fff",padding:"6px",marginBottom:"24px",boxShadow:"0 8px 30px rgba(0,0,0,.3)"}}/>
          <div style={{fontSize:"24px",fontWeight:"800",color:"#fff",marginBottom:"6px",letterSpacing:"-0.5px"}}>SPL Infrastructure Pvt. Ltd.</div>
          <div style={{fontSize:"14px",color:"rgba(255,255,255,.6)",marginBottom:"40px",letterSpacing:"0.05em",textTransform:"uppercase"}}>Daily Progress Report System</div>
          <div style={{background:"rgba(255,255,255,.08)",borderRadius:"14px",padding:"24px 28px",marginBottom:"36px",backdropFilter:"blur(8px)",border:"1px solid rgba(255,255,255,.12)"}}>
            <div style={{fontSize:"20px",color:"rgba(255,255,255,.25)",marginBottom:"10px",fontFamily:"Georgia,serif"}}>"</div>
            <div style={{fontSize:"15px",color:"rgba(255,255,255,.9)",fontStyle:"italic",lineHeight:"1.7",marginBottom:"12px"}}>{q.q}</div>
            {q.a&&<div style={{fontSize:"12px",color:"rgba(255,255,255,.5)",fontWeight:"600"}}>— {q.a}</div>}
          </div>
          <button onClick={()=>setShowLogin(true)} style={{width:"100%",maxWidth:"320px",padding:"17px",borderRadius:"14px",border:"none",background:AM,color:"#fff",cursor:"pointer",fontSize:"17px",fontWeight:"800",boxShadow:"0 4px 20px rgba(245,158,11,.4)",display:"flex",alignItems:"center",justifyContent:"center",gap:"10px",letterSpacing:"0.03em"}}>
            <i className="ti ti-lock-open" style={{fontSize:"20px"}} aria-hidden/>Sign In to Continue
          </button>
          <div style={{fontSize:"12px",color:"rgba(255,255,255,.35)",marginTop:"20px"}}>All data is synced in real-time across devices</div>
          <button onClick={handleGoogleSignOut} style={{marginTop:"14px",background:"none",border:"none",color:"rgba(255,255,255,.55)",fontSize:"12px",fontWeight:"600",cursor:"pointer",textDecoration:"underline",display:"inline-flex",alignItems:"center",gap:"6px"}}><i className="ti ti-logout" aria-hidden/>{gUser?.email?`Not ${gUser.email}? Switch Google account`:"Switch Google account"}</button>
        </div>
      </div>
    );
  }

  // ── NO PROJECT ASSIGNED ──────────────────────────────────────────────────
  if(appView==="noProject"){
    return(
      <div style={{fontFamily:"var(--font-sans)",minHeight:"100vh",background:"#f8fafc",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"24px"}}>
        {toast&&<div style={{position:"fixed",top:"14px",left:"50%",transform:"translateX(-50%)",background:toast.t==="err"?"#fef2f2":"#f0fdf4",color:toast.t==="err"?RD:GN,padding:"12px 22px",borderRadius:"10px",boxShadow:"0 4px 20px rgba(0,0,0,.18)",zIndex:99998,fontSize:"14px",fontWeight:"700",border:`2px solid ${toast.t==="err"?"#fca5a5":"#86efac"}`,display:"flex",alignItems:"center",gap:"8px",maxWidth:"90vw"}}><i className={`ti ti-${toast.t==="err"?"alert-circle":"circle-check"}`} aria-hidden/>{toast.m}</div>}
        <div style={{textAlign:"center",maxWidth:"400px"}}>
          <div style={{fontSize:"56px",marginBottom:"16px"}}>🏗️</div>
          <div style={{fontWeight:"800",fontSize:"20px",color:NV,marginBottom:"8px"}}>No Project Assigned</div>
          <div style={{fontSize:"14px",color:"#6b7280",marginBottom:"24px",lineHeight:"1.6"}}>Your account has not been assigned to a project yet. Please contact your admin to get assigned.</div>
          <div style={{background:"#fffbeb",border:"2px solid #f59e0b",borderRadius:"12px",padding:"16px",marginBottom:"24px",textAlign:"left"}}>
            <div style={{fontWeight:"700",fontSize:"13px",color:"#92400e",marginBottom:"6px"}}>Signed in as:</div>
            <div style={{fontWeight:"800",fontSize:"15px",color:"#111827"}}>{user?.name}</div>
            <RoleB role={user?.role}/>
          </div>
          <button onClick={handleSignOut} style={{padding:"12px 28px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"700"}}>← Sign Out</button>
        </div>
      </div>
    );
  }

  // ── GLOBAL ADMIN / PROJECTS SCREEN ──────────────────────────────────────
  const GLOBAL_VIEWS=["globalAdmin","projects","analytics","reports","performance","auditlog","lateRequests","globalUsers","globalSettings"];
  if(GLOBAL_VIEWS.includes(appView)){
    const pendingReqCount=lateRequests.filter(r=>r.status==="pending").length;
    const ADMIN_TABS=[
      {v:"projects",l:"🏗️ Projects",icon:"ti-building-estate"},
      {v:"analytics",l:"📊 Analytics",icon:"ti-chart-pie"},
      {v:"reports",l:"📄 Reports",icon:"ti-calendar-stats"},
      {v:"performance",l:"⭐ Performance",icon:"ti-star"},
      {v:"auditlog",l:"🕘 Audit Log",icon:"ti-history"},
      {v:"lateRequests",l:"Late Requests",icon:"ti-clock-edit",badge:pendingReqCount},
      {v:"globalUsers",l:"🔐 Users",icon:"ti-users"},
      {v:"globalSettings",l:"⚙️ Settings",icon:"ti-settings"},
    ];
    const activeTabIdx=ADMIN_TABS.findIndex(t=>t.v===appView||(appView==="globalAdmin"&&t.v==="projects"));
    const effectiveView=appView==="globalAdmin"?"projects":appView;

    return(
      <div style={{fontFamily:"var(--font-sans)",background:"#f8fafc",minHeight:"100vh"}}>
        {installBanner}
        {showLogin&&<LoginModal users={users} onLogin={handleLogin} onClose={()=>setShowLogin(false)}/>}
        {toast&&<div style={{position:"fixed",top:"14px",left:"50%",transform:"translateX(-50%)",background:toast.t==="err"?"#fef2f2":"#f0fdf4",color:toast.t==="err"?RD:GN,padding:"12px 22px",borderRadius:"10px",boxShadow:"0 4px 20px rgba(0,0,0,.18)",zIndex:99998,fontSize:"14px",fontWeight:"700",border:`2px solid ${toast.t==="err"?"#fca5a5":"#86efac"}`,display:"flex",alignItems:"center",gap:"8px",maxWidth:"90vw"}}><i className={`ti ti-${toast.t==="err"?"alert-circle":"circle-check"}`} aria-hidden/>{toast.m}</div>}

        <div style={{display:mobile?"block":"flex",alignItems:"stretch"}}>
        {/* ── DESKTOP SIDEBAR ── */}
        {!mobile&&(
          <div style={{width:"236px",flexShrink:0,background:NV,display:"flex",flexDirection:"column",position:"sticky",top:0,height:"100vh",boxShadow:"2px 0 12px rgba(0,0,0,.15)"}}>
            <div style={{display:"flex",alignItems:"center",gap:"10px",padding:"18px 18px 16px"}}>
              <img src={LOGO} alt="SPL" style={{height:"30px",borderRadius:"5px",background:"#fff",padding:"2px"}}/>
              <div><div style={{fontWeight:"800",fontSize:"14px",color:"#fff"}}>SPL DPR</div><div style={{fontSize:"10px",color:"rgba(255,255,255,.55)"}}>Admin Panel</div></div>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:"3px",padding:"6px 12px",flex:1,overflowY:"auto"}}>
              {ADMIN_TABS.map(t=>{const active=effectiveView===t.v;return(
                <button key={t.v} onClick={()=>setAppView(t.v)} style={{display:"flex",alignItems:"center",gap:"11px",width:"100%",padding:"10px 12px",borderRadius:"9px",border:"none",cursor:"pointer",fontSize:"13px",fontWeight:active?"700":"500",textAlign:"left",background:active?"rgba(255,255,255,.14)":"transparent",color:active?"#fff":"#cbd5e1"}}>
                  <i className={`ti ${t.icon}`} style={{fontSize:"16px",flexShrink:0}} aria-hidden/>
                  <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.l}</span>
                  {t.badge>0&&<span style={{background:active?"#fff":RD,color:active?NV:"#fff",fontSize:"10px",fontWeight:"800",borderRadius:"20px",padding:"1px 7px",flexShrink:0}}>{t.badge}</span>}
                </button>
              );})}
            </div>
            <div style={{padding:"12px",borderTop:"1px solid rgba(255,255,255,.1)",display:"flex",flexDirection:"column",gap:"8px"}}>
              {roleKey(user?.role)==="admin"&&<button onClick={async()=>{
                try{flash("⏳ Fetching backup...");
                  const r=await fetch(await authedUrl(RTDB_URL+'/.json'));if(!r.ok)throw new Error('HTTP '+r.status);
                  const data=await r.json();
                  const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
                  const url=URL.createObjectURL(blob);const a=document.createElement('a');
                  a.href=url;a.download='SPL_DPR_Backup_'+new Date().toISOString().slice(0,10)+'.json';
                  a.click();URL.revokeObjectURL(url);flash("✅ Backup downloaded!");
                }catch(e){flash("Backup failed: "+e.message,"err");}
              }} style={{padding:"9px 11px",borderRadius:"8px",border:"1px solid rgba(255,255,255,.2)",background:"transparent",cursor:"pointer",fontSize:"12px",color:"rgba(255,255,255,.75)",display:"flex",alignItems:"center",gap:"7px",fontWeight:"600"}}>
                <i className="ti ti-database-export" style={{fontSize:"14px"}} aria-hidden/>Backup
              </button>}
              <div style={{display:"flex",alignItems:"center",gap:"9px"}}>
                <div style={{width:"32px",height:"32px",borderRadius:"50%",background:"#fef3c7",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"12px",fontWeight:"800",color:"#92400e",flexShrink:0}}>{initials(user?.name)}</div>
                <div style={{flex:1,minWidth:0,lineHeight:"1.25"}}><div style={{fontSize:"12px",fontWeight:"700",color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user?.name}</div><div style={{fontSize:"10px",color:"rgba(255,255,255,.5)"}}>{ROLE_LABELS[roleKey(user?.role)]||user?.role}</div></div>
                <button onClick={handleSignOut} title="Sign out" style={{border:"none",background:"rgba(255,255,255,.08)",color:"#cbd5e1",width:"30px",height:"30px",borderRadius:"8px",cursor:"pointer",fontSize:"14px",flexShrink:0}}><i className="ti ti-logout" aria-hidden/></button>
              </div>
              <div style={{fontSize:"9px",color:"rgba(255,255,255,.3)",textAlign:"center",marginTop:"9px",letterSpacing:".02em"}}>{BUILD_TAG}</div>
            </div>
          </div>
        )}
        <div style={{flex:1,minWidth:0}}>
        {/* ── HEADER ── */}
        <div style={{background:NV,padding:mobile?"10px 14px":"12px 24px",display:"flex",alignItems:"center",justifyContent:"space-between",boxShadow:"0 2px 12px rgba(0,0,0,.25)",position:"sticky",top:0,zIndex:1000,gap:"8px"}}>
          {mobile&&<div style={{display:"flex",alignItems:"center",gap:"8px",flexShrink:0}}>
            <img src={LOGO} alt="SPL" style={{height:"28px",borderRadius:"4px",background:"#fff",padding:"2px"}}/>
          </div>}
          {!mobile&&<div style={{fontWeight:"700",fontSize:"14px",color:"#fff"}}>{ADMIN_TABS.find(t=>t.v===effectiveView)?.l||"Admin"}</div>}
          {mobile&&<div style={{flex:1,textAlign:"center"}}>
            <div style={{fontWeight:"700",fontSize:"14px",color:"#fff"}}>{ADMIN_TABS.find(t=>t.v===effectiveView)?.l||"Admin"}</div>
            <div style={{fontSize:"10px",color:"rgba(255,255,255,.6)"}}>Admin Panel</div>
          </div>}
          {mobile&&roleKey(user?.role)==="admin"&&<button onClick={async()=>{
            try{flash("⏳ Fetching backup...");
              const r=await fetch(await authedUrl(RTDB_URL+'/.json'));if(!r.ok)throw new Error('HTTP '+r.status);
              const data=await r.json();
              const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
              const url=URL.createObjectURL(blob);const a=document.createElement('a');
              a.href=url;a.download='SPL_DPR_Backup_'+new Date().toISOString().slice(0,10)+'.json';
              a.click();URL.revokeObjectURL(url);flash("✅ Backup downloaded!");
            }catch(e){flash("Backup failed: "+e.message,"err");}
          }} title="Download full Firebase backup" style={{padding:"7px 10px",borderRadius:"7px",border:"1px solid rgba(255,255,255,.3)",background:"transparent",cursor:"pointer",fontSize:"12px",color:"rgba(255,255,255,.8)",display:"flex",alignItems:"center",gap:"4px"}}>
            <i className="ti ti-database-export" style={{fontSize:"13px"}} aria-hidden/>
          </button>}
          {mobile&&<button onClick={handleSignOut} style={{padding:"7px 10px",borderRadius:"7px",border:"1px solid rgba(255,255,255,.3)",background:"transparent",cursor:"pointer",fontSize:"12px",color:"rgba(255,255,255,.8)",flexShrink:0}}>Exit</button>}
        </div>
        {mobile&&<div style={{position:"fixed",bottom:0,left:0,right:0,background:"#fff",borderTop:"1px solid #e5e7eb",display:"flex",zIndex:1000,boxShadow:"0 -2px 12px rgba(0,0,0,.1)"}}>
          {ADMIN_TABS.slice(0,4).map(t=>(
            <button key={t.v} onClick={()=>{setAppView(t.v);setShowMoreNav(false);}} style={{flex:1,padding:"10px 4px 12px",border:"none",background:effectiveView===t.v?"#eff6ff":"transparent",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:"3px",position:"relative"}}>
              <i className={`ti ${t.icon}`} style={{fontSize:"22px",color:effectiveView===t.v?NV:"#9ca3af"}} aria-hidden/>
              <span style={{fontSize:"10px",fontWeight:effectiveView===t.v?"700":"500",color:effectiveView===t.v?NV:"#9ca3af"}}>{t.l}</span>
              {t.badge>0&&<span style={{position:"absolute",top:"4px",right:"calc(50% - 16px)",background:RD,color:"#fff",fontSize:"9px",padding:"1px 5px",borderRadius:"10px",fontWeight:"700"}}>{t.badge}</span>}
              {effectiveView===t.v&&<div style={{position:"absolute",top:0,left:"15%",right:"15%",height:"3px",background:NV,borderRadius:"0 0 3px 3px"}}/>}
            </button>
          ))}
          {ADMIN_TABS.length>4&&(()=>{const restActive=ADMIN_TABS.slice(4).some(t=>t.v===effectiveView);const restBadge=ADMIN_TABS.slice(4).reduce((a,t)=>a+(t.badge||0),0);return(
            <button onClick={()=>setShowMoreNav(v=>!v)} style={{flex:1,padding:"10px 4px 12px",border:"none",background:restActive||showMoreNav?"#eff6ff":"transparent",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:"3px",position:"relative"}}>
              <i className="ti ti-dots" style={{fontSize:"22px",color:restActive||showMoreNav?NV:"#9ca3af"}} aria-hidden/>
              <span style={{fontSize:"10px",fontWeight:restActive||showMoreNav?"700":"500",color:restActive||showMoreNav?NV:"#9ca3af"}}>More</span>
              {restBadge>0&&<span style={{position:"absolute",top:"4px",right:"calc(50% - 16px)",background:RD,color:"#fff",fontSize:"9px",padding:"1px 5px",borderRadius:"10px",fontWeight:"700"}}>{restBadge}</span>}
              {(restActive||showMoreNav)&&<div style={{position:"absolute",top:0,left:"15%",right:"15%",height:"3px",background:NV,borderRadius:"0 0 3px 3px"}}/>}
            </button>
          );})()}
        </div>}
        {mobile&&showMoreNav&&(
          <div style={{position:"fixed",inset:0,zIndex:999,background:"rgba(15,23,42,.4)"}} onClick={()=>setShowMoreNav(false)}>
            <div style={{position:"absolute",bottom:"64px",left:"10px",right:"10px",background:"#fff",borderRadius:"14px",boxShadow:"0 -8px 30px rgba(0,0,0,.2)",overflow:"hidden"}} onClick={e=>e.stopPropagation()}>
              {ADMIN_TABS.slice(4).map(t=>(
                <button key={t.v} onClick={()=>{setAppView(t.v);setShowMoreNav(false);}} style={{width:"100%",padding:"14px 18px",border:"none",borderBottom:"1px solid #f1f5f9",background:effectiveView===t.v?"#eff6ff":"#fff",cursor:"pointer",display:"flex",alignItems:"center",gap:"12px",fontSize:"14px",fontWeight:"700",color:effectiveView===t.v?NV:"#374151",textAlign:"left"}}>
                  <i className={`ti ${t.icon}`} style={{fontSize:"18px"}} aria-hidden/>{t.l}
                  {t.badge>0&&<span style={{marginLeft:"auto",background:RD,color:"#fff",fontSize:"11px",padding:"2px 8px",borderRadius:"10px",fontWeight:"800"}}>{t.badge}</span>}
                </button>
              ))}
            </div>
          </div>
        )}
        {/* ── CONTENT ── */}
        <div style={{padding:mobile?"8px":"0",paddingBottom:mobile?"80px":"0",minHeight:"calc(100vh - 60px)"}}>

        {/* PROJECTS */}
        {(effectiveView==="projects")&&(
          <ProjectsScreen projects={projects} user={user} users={users} allSubs={subs} globalLists={globalLists}
            onEnter={p=>{setActiveProject(p);setView("dashboard");setAppView("dashboard");}}
            flash={flash}/>
        )}

        {/* ANALYTICS */}
        {effectiveView==="analytics"&&(
          <AnalyticsScreen projects={projects} users={users} mobile={mobile} flash={flash}/>
        )}

        {/* REPORTS */}
        {effectiveView==="reports"&&(
          <ReportsScreen projects={projects} users={users} mobile={mobile} flash={flash}/>
        )}

        {/* PERFORMANCE */}
        {effectiveView==="performance"&&(
          <PerformanceScreen users={users} projects={projects} mobile={mobile} globalLists={globalLists}/>
        )}

        {/* AUDIT LOG */}
        {effectiveView==="auditlog"&&(
          <AuditLogScreen auditTrail={auditTrail} mobile={mobile}/>
        )}

        {/* LATE-ENTRY REQUESTS */}
        {effectiveView==="lateRequests"&&(
          <LateRequestsScreen lateRequests={lateRequests} decideLateRequest={decideLateRequest} mobile={mobile}/>
        )}

        {/* USERS */}
        {effectiveView==="globalUsers"&&(
          <UsersPanel
            users={users} projects={projects} mobile={mobile}
            newUsr={newUsr} setNewUsr={setNewUsr}
            addUser={addUser} reassignUser={reassignUser}
            saveUserCaps={saveUserCaps} flash={flash}
            globalLists={globalLists} user={user}
          />
        )}

        {/* SETTINGS */}
        {effectiveView==="globalSettings"&&(
          <GlobalSettingsPanel globalLists={globalLists} setGlobalLists={setGlobalLists} flash={flash} mobile={mobile}/>
        )}

        </div>
        </div>
        </div>
      </div>
    );
  }

  return(
    <div style={{fontFamily:"var(--font-sans)",background:"#f8fafc",minHeight:"100vh",paddingBottom:CONTENT_PAD,paddingLeft:mobile?0:"236px"}}>
      {installBanner}
      {showLogin&&<LoginModal users={users} onLogin={handleLogin} onClose={()=>setShowLogin(false)}/>}

      {/* Toast */}
      {toast&&<div style={{position:"fixed",top:"14px",left:"50%",transform:"translateX(-50%)",background:toast.t==="err"?"#fef2f2":"#f0fdf4",color:toast.t==="err"?RD:GN,padding:"12px 22px",borderRadius:"10px",boxShadow:"0 4px 20px rgba(0,0,0,.18)",zIndex:99998,fontSize:"14px",fontWeight:"700",border:`2px solid ${toast.t==="err"?"#fca5a5":"#86efac"}`,display:"flex",alignItems:"center",gap:"8px",maxWidth:"90vw"}}>
        <i className={`ti ti-${toast.t==="err"?"alert-circle":"circle-check"}`} aria-hidden/>{toast.m}
      </div>}

      {/* ─── DESKTOP SIDEBAR ─── */}
      {!mobile&&(
        <div style={{position:"fixed",left:0,top:0,width:"236px",height:"100vh",background:NV,display:"flex",flexDirection:"column",boxShadow:"2px 0 12px rgba(0,0,0,.15)",zIndex:1000}}>
          <div style={{display:"flex",alignItems:"center",gap:"10px",padding:"18px 18px 14px"}}>
            <img src={LOGO} alt="SPL" style={{height:"30px",borderRadius:"5px",background:"#fff",padding:"2px"}}/>
            <div><div style={{fontWeight:"800",fontSize:"14px",color:"#fff"}}>SPL DPR</div><div style={{fontSize:"10px",color:"rgba(255,255,255,.55)"}}>Project workspace</div></div>
          </div>
          <button onClick={()=>safeNav(()=>{
            if(roleKey(user?.role)==="admin"){setActiveProject(null);setView("dashboard");setAppView("globalAdmin");}
            else if(roleKey(user?.role)==="management"){setActiveProject(null);setView("dashboard");setAppView("projects");}
            else{handleSignOut();} // engineers/incharges: sign out instead of showing wrong screen
          })} style={{margin:"0 12px 10px",padding:"9px 11px",borderRadius:"8px",border:"1px solid rgba(255,255,255,.2)",background:"transparent",cursor:"pointer",color:"rgba(255,255,255,.8)",fontSize:"12px",fontWeight:"700",display:"flex",alignItems:"center",gap:"7px"}}>
            <i className="ti ti-arrow-left" style={{fontSize:"14px"}} aria-hidden/>{roleKey(user?.role)==="admin"||roleKey(user?.role)==="management"?"All Projects":"Sign Out"}
          </button>
          <div style={{margin:"0 16px 10px",padding:"10px 12px",background:"rgba(255,255,255,.07)",borderRadius:"9px"}}>
            <div style={{display:"flex",alignItems:"center",gap:"6px"}}>
              <div style={{fontWeight:"800",fontSize:"13px",color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{activeProject.name}</div>
              <div title={!isOnline?"No internet — DPRs queued locally":fbStatus==="ok"?"Live — synced":fbStatus==="error"?"Connection error":"Connecting..."} style={{width:"8px",height:"8px",borderRadius:"50%",background:!isOnline?"#f87171":fbStatus==="ok"?"#4ade80":fbStatus==="error"?"#f87171":"#fbbf24",flexShrink:0}}/>
            </div>
            <div style={{fontSize:"11px",color:"rgba(255,255,255,.6)",marginTop:"2px"}}>{activeProject.code||"DPR System"}{activeProject.location?" — "+activeProject.location:""}</div>
            {!isOnline&&<span style={{display:"inline-block",marginTop:"6px",fontSize:"10px",background:"#f87171",color:"#fff",padding:"1px 7px",borderRadius:"8px",fontWeight:"700"}}>{offlineQueue.length>0?"📵 "+offlineQueue.length+" queued":"📵 Offline"}</span>}
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:"3px",padding:"0 12px",flex:1,overflowY:"auto"}}>
            {allTabs.map(t=>{const active=view===t.id;return(
              <button key={t.id} onClick={()=>safeNav(()=>setView(t.id))} style={{display:"flex",alignItems:"center",gap:"11px",width:"100%",padding:"10px 12px",borderRadius:"9px",border:"none",cursor:"pointer",fontSize:"13px",fontWeight:active?"700":"500",textAlign:"left",background:active?"rgba(255,255,255,.14)":"transparent",color:active?"#fff":"#cbd5e1"}}>
                <i className={`ti ${t.i}`} style={{fontSize:"16px",flexShrink:0}} aria-hidden/>
                <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.l}</span>
                {t.badge>0&&<span style={{background:active?"#fff":RD,color:active?NV:"#fff",fontSize:"10px",fontWeight:"800",borderRadius:"20px",padding:"1px 7px",flexShrink:0}}>{t.badge}</span>}
              </button>
            );})}
          </div>
          <div style={{padding:"12px",borderTop:"1px solid rgba(255,255,255,.1)"}}>
            {user?(
              <div style={{display:"flex",alignItems:"center",gap:"9px"}}>
                <div style={{width:"32px",height:"32px",borderRadius:"50%",background:"#fef3c7",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"12px",fontWeight:"800",color:"#92400e",flexShrink:0}}>{initials(user?.name)}</div>
                <div style={{flex:1,minWidth:0,lineHeight:"1.25"}}><div style={{fontSize:"12px",fontWeight:"700",color:"#fff",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.name}</div><div style={{fontSize:"10px",color:"rgba(255,255,255,.5)"}}>{ROLE_LABELS[roleKey(user?.role)]||user?.role}</div></div>
                <button onClick={handleSignOut} title="Sign out" style={{border:"none",background:"rgba(255,255,255,.08)",color:"#cbd5e1",width:"30px",height:"30px",borderRadius:"8px",cursor:"pointer",fontSize:"14px",flexShrink:0}}><i className="ti ti-logout" aria-hidden/></button>
              </div>
            ):(
              <button onClick={()=>setShowLogin(true)} style={{width:"100%",padding:"10px",borderRadius:"8px",border:"1px solid rgba(255,255,255,.4)",background:"rgba(255,255,255,.12)",cursor:"pointer",fontSize:"13px",color:"#fff",fontWeight:"600",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px"}}><i className="ti ti-lock" style={{fontSize:"14px"}} aria-hidden/>Sign In</button>
            )}
          </div>
        </div>
      )}

      {/* ─── DESKTOP HEADER ─── */}
      {!mobile&&(
        <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 24px",height:"56px",position:"sticky",top:0,zIndex:900,gap:"12px"}}>
          <div style={{fontWeight:"800",fontSize:"16px",color:NV}}>{allTabs.find(t=>t.id===view)?.l||"Dashboard"}</div>
          <div style={{display:"flex",alignItems:"center",gap:"6px"}}>
            {(()=>{const ci=allTabs.findIndex(t=>t.id===view);return(<>
              <button onClick={()=>{if(ci>0)safeNav(()=>setView(allTabs[ci-1].id));}} disabled={ci<=0} title="Previous page" style={{width:"32px",height:"32px",borderRadius:"8px",border:"1px solid #e5e7eb",background:"#fff",color:ci>0?NV:"#cbd5e1",cursor:ci>0?"pointer":"default",fontSize:"18px",display:"flex",alignItems:"center",justifyContent:"center"}}>‹</button>
              <button onClick={()=>{if(ci<allTabs.length-1)safeNav(()=>setView(allTabs[ci+1].id));}} disabled={ci>=allTabs.length-1} title="Next page" style={{width:"32px",height:"32px",borderRadius:"8px",border:"1px solid #e5e7eb",background:"#fff",color:ci<allTabs.length-1?NV:"#cbd5e1",cursor:ci<allTabs.length-1?"pointer":"default",fontSize:"18px",display:"flex",alignItems:"center",justifyContent:"center"}}>›</button>
            </>);})()}
          </div>
        </div>
      )}

      {/* ─── MOBILE TOP BAR ─── */}
      {mobile&&(
        <div style={{background:NV,padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:1000,boxShadow:"0 2px 10px rgba(0,0,0,.25)"}}>
          <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
            <button onClick={()=>safeNav(()=>{
              if(roleKey(user?.role)==="admin"){setActiveProject(null);setView("dashboard");setAppView("globalAdmin");}
              else if(roleKey(user?.role)==="management"){setActiveProject(null);setView("dashboard");setAppView("projects");}
              else{handleSignOut();}
            })} style={{padding:"5px 8px",borderRadius:"6px",border:"1px solid rgba(255,255,255,.3)",background:"transparent",cursor:"pointer",color:"rgba(255,255,255,.8)",fontSize:"12px",display:"flex",alignItems:"center",gap:"3px"}}>
              <i className="ti ti-arrow-left" style={{fontSize:"13px"}} aria-hidden/>
            </button>
            <div><div style={{fontSize:"11px",fontWeight:"700",color:"#fff"}}>{activeProject.name}</div><div style={{fontSize:"10px",color:"rgba(255,255,255,.6)"}}>{activeProject.code||"DPR"}</div></div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
            {pendN>0&&<span style={{background:RD,color:"#fff",fontSize:"11px",padding:"3px 8px",borderRadius:"12px",fontWeight:"700"}}>{pendN} pending</span>}
            {user?(<div style={{display:"flex",alignItems:"center",gap:"6px"}}>
              <RoleB role={user.role}/>
              <button onClick={handleSignOut} style={{padding:"6px 10px",borderRadius:"6px",border:"1px solid rgba(255,255,255,.3)",background:"transparent",cursor:"pointer",fontSize:"12px",color:"rgba(255,255,255,.7)"}}><i className="ti ti-logout" aria-hidden/></button>
            </div>):(
              <button onClick={()=>setShowLogin(true)} style={{padding:"7px 12px",borderRadius:"8px",border:"1px solid rgba(255,255,255,.4)",background:"rgba(255,255,255,.12)",cursor:"pointer",fontSize:"13px",color:"#fff",fontWeight:"600",display:"flex",alignItems:"center",gap:"5px"}}>
                <i className="ti ti-lock" style={{fontSize:"13px"}} aria-hidden/>Sign In
              </button>
            )}
          </div>
        </div>
      )}

      {/* ─── MOBILE BOTTOM TAB BAR ─── */}
      {mobile&&(
        <div style={{position:"fixed",bottom:0,left:0,right:0,background:"#fff",borderTop:"1px solid #e5e7eb",display:"flex",zIndex:1000,boxShadow:"0 -2px 12px rgba(0,0,0,.1)"}}>
          {/* Prev arrow */}
          {(()=>{const ci=allTabs.findIndex(t=>t.id===view);return(
            <button onClick={()=>{if(ci>0)safeNav(()=>setView(allTabs[ci-1].id));}} disabled={ci<=0} style={{padding:"8px 6px",border:"none",background:"transparent",color:ci>0?NV:"#d1d5db",cursor:ci>0?"pointer":"default",fontSize:"20px",flexShrink:0}}>‹</button>
          );})()}
          {allTabs.map(t=>(
            <button key={t.id} onClick={()=>safeNav(()=>setView(t.id))} style={{flex:1,padding:"8px 2px 10px",border:"none",background:view===t.id?"#f0f9ff":"transparent",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:"3px",position:"relative"}}>
              <i className={`ti ${t.i}`} style={{fontSize:"20px",color:view===t.id?NV:"#9ca3af"}} aria-hidden/>
              <span style={{fontSize:"9px",fontWeight:view===t.id?"700":"500",color:view===t.id?NV:"#9ca3af"}}>{t.l}</span>
              {t.badge>0&&<span style={{position:"absolute",top:"4px",right:"calc(50% - 14px)",background:RD,color:"#fff",fontSize:"9px",padding:"1px 5px",borderRadius:"10px",fontWeight:"700"}}>{t.badge}</span>}
              {view===t.id&&<div style={{position:"absolute",top:0,left:"20%",right:"20%",height:"2px",background:NV,borderRadius:"0 0 2px 2px"}}/>}
            </button>
          ))}
          {/* Next arrow */}
          {(()=>{const ci=allTabs.findIndex(t=>t.id===view);return(
            <button onClick={()=>{if(ci<allTabs.length-1)safeNav(()=>setView(allTabs[ci+1].id));}} disabled={ci>=allTabs.length-1} style={{padding:"8px 6px",border:"none",background:"transparent",color:ci<allTabs.length-1?NV:"#d1d5db",cursor:ci<allTabs.length-1?"pointer":"default",fontSize:"20px",flexShrink:0}}>›</button>
          );})()}
        </div>
      )}

      <div style={{padding:PAD,maxWidth:"1300px",margin:"0 auto"}}>

      {/* ═══ DASHBOARD ═══ */}
      {view==="dashboard"&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"14px",flexWrap:"wrap"}}>
            <span style={{fontWeight:"700",color:"#374151",fontSize:"14px"}}>DPR for:</span>
            <input type="date" value={dashDate} onChange={e=>setDashDate(e.target.value)} style={{fontSize:"14px",padding:"8px 10px",borderRadius:"8px",border:"1.5px solid #d1d5db"}}/>
            <button onClick={()=>setDashDate(new Date(Date.now()-86400000).toISOString().slice(0,10))} style={{fontSize:"13px",padding:"8px 14px",borderRadius:"8px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",color:"#6b7280",fontWeight:"600"}}>{mobile?"←":"Yesterday"}</button>
            <button onClick={()=>setDashDate(new Date().toISOString().slice(0,10))} style={{fontSize:"13px",padding:"8px 14px",borderRadius:"8px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",color:"#6b7280",fontWeight:"600"}}>Today</button>
            {caps.download&&<button onClick={()=>doExcel(subs,engineers,dashDate,activeProject?.name)} style={{marginLeft:"auto",padding:mobile?"8px":"8px 16px",borderRadius:"8px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700",display:"flex",alignItems:"center",gap:"6px"}} title="Download Excel">
              <i className="ti ti-file-spreadsheet" style={{fontSize:"16px"}} aria-hidden/>
              {!mobile&&"Download Excel"}
            </button>}
            <button onClick={()=>doPrintDPR(dayS,engineers,dashDate,activeProject?.name||"Project")} style={{marginLeft:!caps.download?"auto":"0",padding:mobile?"8px":"8px 16px",borderRadius:"8px",border:"1px solid #e5e7eb",background:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700",display:"flex",alignItems:"center",gap:"6px",color:"#374151"}} title="Print DPR">
              <i className="ti ti-printer" style={{fontSize:"16px"}} aria-hidden/>
              {!mobile&&"Print DPR"}
            </button>
            <button onClick={()=>setShareOpen(true)} style={{padding:mobile?"8px":"8px 16px",borderRadius:"8px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700",display:"flex",alignItems:"center",gap:"6px",color:"#374151"}} title="Share">
              <i className="ti ti-share" style={{fontSize:"16px"}} aria-hidden/>
              {!mobile&&"Share"}
            </button>
          </div>
          {/* Metric cards */}
          <div style={{display:"grid",gridTemplateColumns:mobile?"repeat(2,1fr)":"repeat(auto-fit,minmax(120px,1fr))",gap:"10px",marginBottom:"16px"}}>
            {[{l:"Present",v:presentN.size,c:GN,bg:"#f0fdf4",i:"ti-user-check"},{l:"Absent",v:engineers.length-presentN.size,c:RD,bg:"#fef2f2",i:"ti-user-off"},{l:"Day Shift",v:[...new Set(dayS.filter(s=>s.shift==="Day").map(s=>s.engineer))].length,c:"#d97706",bg:"#fffbeb",i:"ti-sun-high"},{l:"Night Shift",v:[...new Set(dayS.filter(s=>s.shift==="Night").map(s=>s.engineer))].length,c:"#2563eb",bg:"#eff6ff",i:"ti-moon"},{l:"Mat.Moves",v:flatM.length,c:"#0f766e",bg:"#f0fdfa",i:"ti-package"},{l:"Pending ✓",v:pendN,c:RD,bg:"#fef2f2",i:"ti-clock"}].map(m=>(
              <div key={m.l} style={{background:m.bg,borderRadius:"12px",padding:"12px 14px",border:`1px solid ${m.c}20`}}>
                <div style={{display:"flex",alignItems:"center",gap:"6px",marginBottom:"7px"}}><i className={`ti ${m.i}`} style={{fontSize:"16px",color:m.c}} aria-hidden/><span style={{fontSize:"11px",color:"#6b7280",fontWeight:"600",lineHeight:"1.3"}}>{m.l}</span></div>
                <div style={{fontSize:"28px",fontWeight:"800",color:m.c}}>{m.v}</div>
              </div>
            ))}
          </div>
          {/* Greeting cards */}
          {roleKey(user?.role)==="engineer"&&(()=>{const engRec=engineers.find(x=>x.name===user.name||x.id===user.id);return(<div style={{background:`linear-gradient(135deg,${NV},#2d5a9e)`,borderRadius:"14px",padding:"16px 20px",marginBottom:"14px",color:"#fff",display:"flex",alignItems:"center",gap:"14px",flexWrap:"wrap"}}><Av name={user.name} sz={44}/><div style={{flex:1}}><div style={{fontWeight:"800",fontSize:"17px",marginBottom:"3px"}}>Welcome, {user.name.split(" ")[0]}!</div><div style={{fontSize:"13px",color:"rgba(255,255,255,.8)",display:"flex",gap:"12px",flexWrap:"wrap"}}>{engRec?.incharge?<span>👤 Incharge: <strong>{engRec.incharge}</strong></span>:<span style={{color:"#fbbf24"}}>⚠️ No incharge set</span>}{engRec?.dept&&<span>🏗️ {engRec.dept}</span>}{engRec?.designation&&<span>🏷️ {engRec.designation}</span>}</div></div><button onClick={openForm} style={{padding:"10px 18px",borderRadius:"10px",border:"2px solid rgba(255,255,255,.4)",background:"rgba(255,255,255,.15)",cursor:"pointer",fontSize:"14px",color:"#fff",fontWeight:"700",display:"flex",alignItems:"center",gap:"6px",flexShrink:0}}><i className="ti ti-pencil-plus" aria-hidden/>Fill DPR</button></div>);})()}
          {roleKey(user?.role)==="incharge"&&(()=>{const myEngs=engineers.filter(e=>e.incharge===user.name);const todayStr=new Date().toISOString().slice(0,10);const pendingCount=myEngs.filter(e=>subs.some(s=>s.engineer===e.name&&!s.approved&&!s.needsRevision)).length;return(<div style={{background:`linear-gradient(135deg,#0f766e,#0d9488)`,borderRadius:"14px",padding:"16px 20px",marginBottom:"14px",color:"#fff"}}><div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"10px",flexWrap:"wrap"}}><Av name={user.name} sz={40}/><div style={{flex:1}}><div style={{fontWeight:"800",fontSize:"17px"}}>Welcome, {user.name.split(" ")[0]}!</div><div style={{fontSize:"12px",color:"rgba(255,255,255,.75)"}}>Incharge — {myEngs.length} engineer{myEngs.length!==1?"s":""}</div></div>{pendingCount>0&&<div style={{background:"#fbbf24",color:"#92400e",borderRadius:"10px",padding:"6px 14px",fontWeight:"700",fontSize:"13px"}}>{pendingCount} pending</div>}</div><div style={{display:"flex",gap:"6px",flexWrap:"wrap"}}>{myEngs.map(e=>{const submitted=subs.some(s=>s.engineer===e.name&&s.date===todayStr);const approved=subs.some(s=>s.engineer===e.name&&s.date===todayStr&&s.approved);return<div key={e.id} style={{background:"rgba(255,255,255,.15)",borderRadius:"7px",padding:"5px 10px",fontSize:"12px",fontWeight:"600"}}>{approved?"✅":submitted?"⏳":"❌"} {e.name}</div>;})}</div></div>);})()}
          {!user&&<div style={{background:"#eff6ff",border:"2px solid #93c5fd",borderRadius:"12px",padding:"14px 16px",marginBottom:"16px",display:"flex",alignItems:"center",gap:"12px",flexWrap:"wrap"}}>
            <div style={{flex:1}}><div style={{fontWeight:"700",fontSize:"14px",color:"#1e40af"}}>Ready to submit today's work?</div></div>
            <button onClick={openForm} style={{padding:"10px 22px",borderRadius:"10px",border:"none",background:AM,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"700",display:"flex",alignItems:"center",gap:"6px"}}><i className="ti ti-pencil-plus" aria-hidden/>Fill DPR</button>
          </div>}
          {myRevisions.length>0&&<div style={{background:"#fef3c7",border:"2px solid #f59e0b",borderRadius:"12px",padding:"14px 16px",marginBottom:"16px"}}>
            <div style={{fontWeight:"700",fontSize:"14px",color:"#92400e",marginBottom:"8px"}}>⚠️ You have {myRevisions.length} DPR(s) marked for revision</div>
            {myRevisions.map(s=>(
              <div key={s.id} style={{background:"#fff",borderRadius:"8px",padding:"10px 14px",marginBottom:"6px",display:"flex",alignItems:"center",gap:"10px",flexWrap:"wrap"}}>
                <div style={{flex:1}}><div style={{fontWeight:"700",fontSize:"13px"}}>{s.date} — {s.shift} Shift</div><div style={{fontSize:"12px",color:RD,marginTop:"3px"}}>Revision note: {s.revisionNote||"Please check and re-submit"}</div></div>
                <button onClick={()=>loadSubForEdit(s)} style={{padding:"8px 16px",borderRadius:"8px",border:"none",background:AM,color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700"}}>✏️ Edit & Resubmit</button>
              </div>
            ))}
          </div>}
          {myLateRequests.filter(r=>r.status!=="approved").length>0&&<div style={{background:"#eff6ff",border:"2px solid #93c5fd",borderRadius:"12px",padding:"14px 16px",marginBottom:"16px"}}>
            <div style={{fontWeight:"700",fontSize:"14px",color:"#1e40af",marginBottom:"8px"}}><i className="ti ti-clock-edit" aria-hidden/> Your late-entry request{myLateRequests.filter(r=>r.status!=="approved").length!==1?"s":""}</div>
            {myLateRequests.filter(r=>r.status!=="approved").map(r=>(
              <div key={r.id} style={{background:"#fff",borderRadius:"8px",padding:"10px 14px",marginBottom:"6px",display:"flex",alignItems:"center",gap:"10px",flexWrap:"wrap"}}>
                <div style={{flex:1}}><div style={{fontWeight:"700",fontSize:"13px"}}>{r.date}</div><div style={{fontSize:"12px",color:"#6b7280",marginTop:"3px"}}>{r.reason}</div></div>
                <span style={{fontSize:"12px",fontWeight:"800",color:r.status==="pending"?"#92400e":RD,background:r.status==="pending"?"#fef3c7":"#fee2e2",borderRadius:"8px",padding:"5px 11px"}}>{r.status==="pending"?"⏳ Pending":"✗ Rejected"}</span>
              </div>
            ))}
          </div>}
          {/* Attendance */}
          <Card style={{marginBottom:"14px"}}>
            <div style={SH}><span>Engineer Attendance — {new Date(dashDate+"T12:00:00").toLocaleDateString("en-IN",{weekday:"short",day:"2-digit",month:"short",year:"numeric"})}</span></div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:mobile?"13px":"12px"}}>
                <thead><tr style={{background:"#f8fafc"}}>{(mobile?["Engineer","DPR Status","Shift"]:["Engineer","Dept","Incharge","DPR Status","Shift","Activities","Mat.Moves","Saved","Approved By"]).map(h=><th key={h} style={{padding:"10px",textAlign:"left",fontWeight:"700",color:"#6b7280",fontSize:"11px",borderBottom:"1px solid #e5e7eb",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>
                  {engineers.map(e=>{
                    const ss=dayS.filter(s=>s.engineer===e.name);
                    const pres=ss.length>0;
                    const shifts=[...new Set(ss.map(s=>s.shift))];
                    const actCnt=ss.reduce((a,s)=>a+(s.activities||[]).length,0);
                    const matCnt=ss.reduce((a,s)=>a+(s.matTxs||[]).length,0);
                    const lastT=ss.length>0?new Date(Math.max(...ss.map(s=>new Date(s.submittedAt)))).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"}):"—";
                    const apvBy=ss.filter(s=>s.approved).map(s=>s.approvedBy).filter(Boolean)[0]||null;
                    const hasRevision=ss.some(s=>s.needsRevision);
                    const allApproved=pres&&ss.every(s=>s.approved);
                    const someApproved=pres&&ss.some(s=>s.approved)&&!allApproved;
                    const dprStatus=!pres
                      ?{label:"❌ Not Submitted",color:RD,bg:"#fef2f2"}
                      :hasRevision
                      ?{label:"↩ Revision Sent",color:"#92400e",bg:"#fef3c7"}
                      :allApproved
                      ?{label:"✅ Approved",color:GN,bg:"#f0fdf4"}
                      :someApproved
                      ?{label:"🟡 Part Approved",color:"#d97706",bg:"#fffbeb"}
                      :{label:"⏳ Pending Approval",color:"#d97706",bg:"#fffbeb"};
                    return(
                      <tr key={e.id} style={{borderBottom:"1px solid #f3f4f6"}}>
                        <td style={{padding:"10px"}}><div style={{display:"flex",alignItems:"center",gap:"8px"}}><Av name={e.name} sz={mobile?32:28}/><div><div style={{fontWeight:"700"}}>{e.name}</div>{ss.length>1&&<span style={{background:"#818cf8",color:"#fff",fontSize:"10px",padding:"1px 6px",borderRadius:"8px",fontWeight:"700"}}>{ss.length} entries</span>}{mobile&&<div style={{fontSize:"10px",color:"#6b7280"}}>{e.dept}</div>}</div></div></td>
                        {!mobile&&<td style={{padding:"10px"}}><DeptB dept={e.dept}/></td>}
                        {!mobile&&<td style={{padding:"10px",color:"#6b7280",fontSize:"11px"}}>{e.incharge}</td>}
                        <td style={{padding:"10px"}}><span style={{background:dprStatus.bg,color:dprStatus.color,padding:"4px 10px",borderRadius:"8px",fontSize:"11px",fontWeight:"700",whiteSpace:"nowrap"}}>{dprStatus.label}</span></td>
                        <td style={{padding:"10px"}}><div style={{display:"flex",gap:"3px",flexWrap:"wrap"}}>{shifts.map(s=><Pill key={s} label={s} color={s==="Day"?"#d97706":"#2563eb"} bg={s==="Day"?"#fffbeb":"#eff6ff"}/>)}{!pres&&<span style={{color:"#9ca3af",fontSize:"11px"}}>—</span>}</div></td>
                        {!mobile&&<td style={{padding:"10px",textAlign:"center",fontWeight:"700",color:actCnt>0?PU:"#9ca3af"}}>{pres?actCnt:"—"}</td>}
                        {!mobile&&<td style={{padding:"10px",textAlign:"center",fontWeight:"700",color:matCnt>0?"#0f766e":"#9ca3af"}}>{pres?matCnt:"—"}</td>}
                        {!mobile&&<td style={{padding:"10px",color:"#6b7280",fontSize:"11px"}}>{lastT}</td>}
                        <td style={{padding:"10px"}}>{apvBy?<Pill label={"✓ "+apvBy} color={GN} bg="#f0fdf4"/>:pres?<Pill label="Pending" color="#d97706" bg="#fffbeb"/>:<span style={{color:"#9ca3af",fontSize:"11px"}}>—</span>}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </Card>
          
          {/* My DPR entries — shows today's subs for logged-in user */}
          {user&&(()=>{
            const todayStr=new Date().toISOString().slice(0,10);
            const mySubs=subs.filter(s=>s.engineer===user.name&&s.date===todayStr);
            if(mySubs.length===0)return(
              <Card style={{marginBottom:"14px",border:`2px dashed ${NV}30`,textAlign:"center",padding:"2rem"}}>
                <i className="ti ti-pencil-plus" style={{fontSize:"32px",color:NV,display:"block",marginBottom:"8px"}} aria-hidden/>
                <div style={{fontWeight:"700",fontSize:"15px",color:NV,marginBottom:"6px"}}>No DPR filled for today yet</div>
                <button onClick={openForm} style={{padding:"10px 24px",borderRadius:"10px",border:"none",background:AM,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"700",marginTop:"8px"}}>Fill Today's DPR →</button>
              </Card>
            );
            return(
              <Card style={{marginBottom:"14px",border:`2px solid ${NV}25`,borderTop:`4px solid ${NV}`}}>
                <div style={SH}>
                  <span>Your DPR — Today ({todayStr}){mySubs.length>1&&<span style={{background:"#818cf8",color:"#fff",borderRadius:"8px",padding:"1px 8px",fontSize:"11px",marginLeft:"8px"}}>{mySubs.length} submissions</span>}</span>
                  <button onClick={openForm} style={{padding:"6px 14px",borderRadius:"7px",border:"none",background:AM,color:"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"700"}}>+ Add More</button>
                </div>
                {mySubs.length===0&&<div style={{fontSize:"13px",color:"#9ca3af",padding:"8px 0"}}>No submissions for today yet.</div>}
                {mySubs.map((s)=>(
                  <div key={s.id} style={{marginBottom:"12px",padding:"12px",borderRadius:"10px",border:`1px solid ${s.needsRevision?"#f59e0b":s.approved?"#86efac":"#e5e7eb"}`,background:s.needsRevision?"#fffbeb":s.approved?"#f0fdf4":"#f8fafc"}}>
                    <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"8px",flexWrap:"wrap"}}>
                      <Pill label={s.shift+" Shift"} color={s.shift==="Day"?"#d97706":"#2563eb"} bg={s.shift==="Day"?"#fffbeb":"#eff6ff"}/>
                      <span style={{fontSize:"11px",color:"#9ca3af",display:"flex",alignItems:"center",gap:"3px"}}>
                        <i className="ti ti-clock" style={{fontSize:"11px"}} aria-hidden/>
                        Saved {new Date(s.submittedAt).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"})}
                      </span>
                      {s.weather&&<span style={{fontSize:"11px",color:"#6b7280"}}>{s.weather}</span>}
                      {s.approved&&<><Pill label={"✓ "+s.approvedBy} color={GN} bg="#f0fdf4"/>{s.approvedAt&&<span style={{fontSize:"11px",color:"#6b7280",marginLeft:"6px"}}>{new Date(s.approvedAt).toLocaleDateString("en-IN",{day:"2-digit",month:"short"})} at {new Date(s.approvedAt).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"})}</span>}</>}
                      {s.needsRevision&&<><Pill label="↩ Needs Revision" color="#92400e" bg="#fef3c7"/><button onClick={()=>loadSubForEdit(s)} style={{padding:"5px 12px",borderRadius:"6px",border:"none",background:AM,color:"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"700"}}>✏️ Edit</button></>}
                      {!s.approved&&!s.needsRevision&&<Pill label="⏳ Pending" color="#d97706" bg="#fffbeb"/>}
                    </div>
                    {(s.activities||[]).length>0&&<div>
                      <div style={{fontSize:"11px",fontWeight:"700",color:"#374151",marginBottom:"5px",textTransform:"uppercase",letterSpacing:"0.04em"}}>Activities ({s.activities.length})</div>
                      {(s.activities||[]).map((a,i)=>(
                        <div key={i} style={{fontSize:"12px",padding:"5px 10px",borderRadius:"6px",background:"#fff",border:"1px solid #e5e7eb",marginBottom:"3px",display:"flex",gap:"8px",flexWrap:"wrap",alignItems:"center"}}>
                          <span style={{fontWeight:"700",color:NV}}>{i+1}.</span>
                          <span style={{fontWeight:"600"}}>{a.actType||(a.actCustom||"—")}</span>
                          {a.desc&&<span style={{color:"#6b7280"}}>— {a.desc}</span>}
                          {a.chFrom&&<span style={{color:"#9ca3af",fontFamily:"monospace",fontSize:"11px"}}>CH {a.chFrom}→{a.chTo||"?"}</span>}
                          {a.theoQty&&<span style={{color:PU,fontWeight:"700"}}>{a.theoQty} {a.unit}</span>}
                        </div>
                      ))}
                    </div>}
                    {(s.matTxs||[]).length>0&&<div style={{marginTop:"6px"}}>
                      <div style={{fontSize:"11px",fontWeight:"700",color:"#374151",marginBottom:"5px",textTransform:"uppercase",letterSpacing:"0.04em"}}>Material Moves ({s.matTxs.length})</div>
                      {(s.matTxs||[]).map((m,i)=>(
                        <div key={i} style={{fontSize:"12px",padding:"5px 10px",borderRadius:"6px",background:"#fff",border:"1px solid #e5e7eb",marginBottom:"3px",display:"flex",gap:"8px",flexWrap:"wrap",alignItems:"center"}}>
                          <span style={{fontWeight:"700",color:"#0f766e"}}>{i+1}.</span>
                          <span style={{fontWeight:"600"}}>{m.material==="Other"?m.matCustom:m.material}</span>
                          <Pill label={m.txType} color="#0f766e" bg="#f0fdfa"/>
                          {(m.recvQty||m.sendQty)&&<span style={{color:"#6b7280"}}>{m.recvQty||m.sendQty} {m.recvUnit||m.sendUnit}</span>}
                        </div>
                      ))}
                    </div>}
                    {s.difficulty&&s.difficulty!=="None"&&<div style={{marginTop:"6px",fontSize:"12px",color:"#92400e",background:"#fef3c7",padding:"5px 10px",borderRadius:"6px"}}>⚠️ {s.difficulty}</div>}
                  </div>
                ))}
              </Card>
            );
          })()}
        </div>
      )}

      {/* ═══ FILL DPR ═══ */}
      {view==="form"&&(
        <div>
          {!user&&(
            <Card style={{textAlign:"center",padding:"3rem"}}>
              <i className="ti ti-lock" style={{fontSize:"48px",color:NV,display:"block",marginBottom:"16px"}} aria-hidden/>
              <div style={{fontWeight:"800",fontSize:"20px",color:NV,marginBottom:"8px"}}>Sign In Required</div>
              <div style={{fontSize:"14px",color:"#6b7280",marginBottom:"24px",maxWidth:"320px",margin:"0 auto 24px"}}>You must be signed in to fill a DPR. This ensures every entry is linked to the correct engineer.</div>
              <button onClick={()=>setShowLogin(true)} style={{padding:"14px 32px",borderRadius:"12px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"15px",fontWeight:"800",display:"inline-flex",alignItems:"center",gap:"8px"}}><i className="ti ti-lock" aria-hidden/>Sign In to Continue</button>
            </Card>
          )}
          {user&&<>
          {/* Step indicator — sticky on mobile so site engineers always see where they are */}
          <div style={{display:"flex",gap:mobile?"6px":"4px",marginBottom:"16px",overflowX:"auto",paddingBottom:"4px",...(mobile?{position:"sticky",top:"46px",zIndex:500,background:"#f8fafc",paddingTop:"8px",margin:"-8px -2px 12px",paddingLeft:"2px",paddingRight:"2px",boxShadow:"0 6px 8px -8px rgba(0,0,0,.25)"}:{})}}>
            {[{t:mobile?"Details":"1 — Your Details"},{t:mobile?"Materials":"2 — Material Moves"},{t:mobile?"Work":"3 — Work Activities"},{t:mobile?"Submit":"4 — Review & Submit"}].map((s,idx)=>(
              <button key={idx} onClick={()=>setStep(idx)} style={{display:"flex",alignItems:"center",gap:"6px",padding:mobile?"11px 13px":"8px 16px",borderRadius:"10px",cursor:"pointer",flexShrink:0,flex:mobile?1:undefined,justifyContent:mobile?"center":undefined,border:`2px solid ${idx===step?AM:idx<step?GN:"#e5e7eb"}`,background:idx===step?"#fffbeb":idx<step?"#f0fdf4":"#fff"}}>
                <div style={{width:"22px",height:"22px",borderRadius:"50%",background:idx===step?AM:idx<step?GN:"#d1d5db",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{idx<step?<span style={{fontSize:"12px",color:"#fff",fontWeight:"800"}}>✓</span>:<span style={{fontSize:"11px",color:idx===step?"#fff":"#9ca3af",fontWeight:"800"}}>{idx+1}</span>}</div>
                <span style={{fontSize:mobile?"12px":"12px",fontWeight:"700",color:idx===step?"#92400e":idx<step?GN:"#9ca3af",whiteSpace:"nowrap"}}>{s.t}</span>
              </button>
            ))}
          </div>

          {/* Step 1: Details */}
          {step===0&&(
            <Card style={{borderTop:"4px solid "+NV}}>
              <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"22px",paddingBottom:"14px",borderBottom:"1px solid #e5e7eb",flexWrap:"wrap"}}>
                <div style={{background:NV,padding:"10px",borderRadius:"10px",flexShrink:0,display:"flex"}}><i className="ti ti-id-badge-2" style={{fontSize:"22px",color:"#fff",display:"block"}} aria-hidden/></div>
                <div style={{flex:1}}><div style={{fontWeight:"800",fontSize:"17px",color:NV}}>Your Details</div><div style={{fontSize:"13px",color:"#6b7280",marginTop:"2px"}}>Enter your information before adding work.</div></div>
              </div>
              <Grid cols={mobile?"1fr":"1fr 1fr"}>
                <F lbl="Date" req>
                  <input type="date" value={hdr.date} onChange={e=>{updHdr("date",e.target.value);setStepErrors(p=>({...p,date:undefined}));}} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:`1.5px solid ${stepErrors.date?"#dc2626":"#d1d5db"}`,fontSize:"15px",background:stepErrors.date?"#fef2f2":"#fff"}}/>
                  {stepErrors.date&&<div style={{color:"#dc2626",fontSize:"11px",marginTop:"3px",fontWeight:"600"}}>Required — select a date</div>}
                </F>
                <F lbl="Shift" req><Seg opts={["Day","Night"]} val={hdr.shift} onChange={v=>updHdr("shift",v)}/></F>
                <F lbl="Your Name" req>
                  {roleKey(user?.role)==="engineer"
                    ?<div style={{padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #86efac",background:"#f0fdf4",fontSize:"15px",fontWeight:"700",color:"#14532d",display:"flex",alignItems:"center",gap:"8px"}}>
                        <i className="ti ti-lock" style={{fontSize:"14px",color:GN}} aria-hidden/>{user.name}
                        <span style={{fontSize:"11px",color:"#6b7280",fontWeight:"400",marginLeft:"auto"}}>Signed in</span>
                      </div>
                    :<select value={hdr.engineer} onChange={e=>updHdr("engineer",e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px",color:"#111827",background:"#fff"}}>
                      <option value="">— Select Engineer Name —</option>
                      {engineers.map(e=><option key={e.id}>{e.name}</option>)}
                      <option value="__other">Other (not in list)</option>
                    </select>
                  }
                </F>
                <F lbl="Department">
                  <select value={hdr.dept} onChange={e=>updHdr("dept",e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px",color:hdr.dept?"#111827":"#6b7280",background:"#fff"}}>
                    <option value="">— Select Department —</option>
                    {(lists.depts||DEPTS).map(d=><option key={d} value={d}>{d}</option>)}
                  </select>
                </F>
                {hdr.engineer==="__other"&&<F lbl="Your Name (type here)" col={2}><Inp value={hdr.engCustom} onChange={e=>updHdr("engCustom",e.target.value)} placeholder="Type your full name"/></F>}
                <F lbl="Incharge / HOD" col={mobile?undefined:2} req={globalLists.requiredFields?.incharge}><Inp value={hdr.incharge} onChange={e=>{updHdr("incharge",e.target.value);setStepErrors(p=>({...p,incharge:undefined}));}} placeholder="Auto-filled — or type here"/>{stepErrors.incharge&&<div style={{color:"#dc2626",fontSize:"11px",marginTop:"3px",fontWeight:"600"}}>{stepErrors.incharge}</div>}</F>
                <F lbl="Weather Condition" req={globalLists.requiredFields?.weather}><Sel val={hdr.weather||"☀️ Clear"} onChange={v=>updHdr("weather",v)} opts={WEATHER_OPTS}/>{stepErrors.weather&&<div style={{color:"#dc2626",fontSize:"11px",marginTop:"3px",fontWeight:"600"}}>{stepErrors.weather}</div>}</F>
                <F lbl="Issues / Problems Faced (describe briefly)" req={globalLists.requiredFields?.difficulty}><Inp value={hdr.difficulty||""} onChange={e=>{updHdr("difficulty",e.target.value);setStepErrors(p=>({...p,difficulty:undefined}));}} placeholder="e.g. Vehicle breakdown at CH 39+200, work stopped 2hrs"/>{stepErrors.difficulty&&<div style={{color:"#dc2626",fontSize:"11px",marginTop:"3px",fontWeight:"600"}}>{stepErrors.difficulty}</div>}</F>
              </Grid>
            </Card>
          )}

          {/* Step 2: Material */}
          {step===1&&(
            <div>
              <div style={{background:"#f0fdfa",border:"2px solid #5eead4",borderRadius:"12px",padding:"14px 16px",marginBottom:"14px"}}>
                <div style={{fontWeight:"700",fontSize:"15px",color:"#0f766e",marginBottom:"4px"}}>📦 Material Movements for the Day</div>
                <div style={{fontSize:"13px",color:"#115e59"}}>Record any material received at site or transferred between chainages. Leave empty if nothing today.</div>
              </div>
              {matTxs.map((tx,i)=>(
                <Card key={tx._k} style={{marginBottom:"16px",borderLeft:`4px solid ${editMatIdx===i?"#f59e0b":"#14b8a6"}`,boxShadow:"0 2px 8px rgba(0,0,0,.06)"}}>
                  {editMatIdx===i?(
                    <MatTxNew editData={tx} onEditSave={updated=>{setMatTxs(p=>p.map((t,idx)=>idx===i?{...updated,_k:t._k}:t));setEditMatIdx(null);flash("✓ Updated");}} onEditCancel={()=>setEditMatIdx(null)} lists={lists}/>
                  ):(
                    <>
                      <div style={{display:"flex",alignItems:"center",gap:"10px",marginBottom:"8px"}}>
                        <div style={{width:"26px",height:"26px",borderRadius:"50%",background:"#14b8a6",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"12px",fontWeight:"800",color:"#fff",flexShrink:0}}>{i+1}</div>
                        <div style={{flex:1,fontWeight:"700",fontSize:"14px",color:"#0f766e"}}>{tx.material==="Other"?tx.matCustom:tx.material||"material"} — {tx.txType}</div>
                        <button onClick={()=>setEditMatIdx(i)} style={{padding:"6px 12px",borderRadius:"6px",border:"1px solid #93c5fd",background:"#eff6ff",color:"#1d4ed8",cursor:"pointer",fontSize:"12px",fontWeight:"600"}}>✏️ Edit</button>
                        <button onClick={()=>setMatTxs(p=>p.filter(t=>t._k!==tx._k))} style={{padding:"6px 10px",borderRadius:"6px",border:"1px solid #fca5a5",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"13px"}}>✕</button>
                      </div>
                      <div style={{fontSize:"13px",color:"#374151",display:"flex",gap:"10px",flexWrap:"wrap"}}>
                        {tx.recvCH&&<span>📥 CH: <strong>{tx.recvCH}</strong></span>}
                        {tx.source&&<span>From: {tx.source}</span>}
                        {tx.recvQty&&<span>Qty: <strong>{tx.recvQty} {tx.recvUnit}</strong></span>}
                        {tx.recvLoads&&<span>{tx.recvLoads} loads</span>}
                        {tx.sendFromCH&&<span>📤 <strong>{tx.sendFromCH}</strong> → {tx.sendToCH}</span>}
                        {tx.sendQty&&<span>Qty: <strong>{tx.sendQty} {tx.sendUnit}</strong></span>}
                        {(tx.assets||[]).length>0&&<span>⚙️ {tx.assets.length} machine(s)</span>}
                      </div>
                    </>
                  )}
                </Card>
              ))}
              {matTxs.length>0&&<div style={{borderTop:"2px dashed #ccfbf1",margin:"4px 0 16px",textAlign:"center"}}><span style={{background:"#fff",padding:"0 10px",fontSize:"11px",color:"#6b7280",fontWeight:"600",position:"relative",top:"-10px"}}>+ Add Another Entry</span></div>}
              <MatTxNew onAdd={tx=>{setMatTxs(p=>[...p,tx]);setHasUnsavedForm(false);flash("✓ Material entry added");}} lists={lists} onDirty={()=>setHasUnsavedForm(true)} requiredFields={globalLists.requiredFields}/>
            </div>
          )}

          {/* Step 3: Work Activities */}
          {step===2&&(
            <div>
              {acts.length>0&&(
                <Card style={{marginBottom:"14px",borderTop:"3px solid "+GN}}>
                  <div style={{fontWeight:"700",fontSize:"14px",color:GN,marginBottom:"12px"}}>✅ Activities Added — {acts.length}</div>
                  {acts.map((a,i)=>(
                    <div key={a._k} style={{display:"flex",alignItems:"flex-start",gap:"10px",padding:"14px",borderRadius:"12px",border:`2px solid ${editActIdx===i?"#f59e0b":"#bbf7d0"}`,background:editActIdx===i?"#fffbeb":"#f0fdf4",marginBottom:"14px",boxShadow:"0 2px 8px rgba(0,0,0,.05)"}}>
                      <div style={{width:"26px",height:"26px",borderRadius:"50%",background:editActIdx===i?AM:GN,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"11px",color:"#fff",fontWeight:"800",flexShrink:0,marginTop:"2px"}}>{i+1}</div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:"14px",fontWeight:"700",color:"#14532d",marginBottom:"3px"}}>{a.actType||(a.actCustom||"—")}{a.desc?` — ${a.desc}`:""}</div>
                        <div style={{fontSize:"12px",color:"#166534",fontFamily:"monospace"}}>{[a.chFrom&&`CH ${a.chFrom}→${a.chTo||"?"}`,a.side,a.theoQty&&`${a.theoQty} ${a.unit}`,(a.assets||[]).length>0&&`${a.assets.length} machine(s)`,(a.contractors||[]).length>0&&a.contractors.map(c=>c.name).join(", ")].filter(Boolean).join("  ·  ")}</div>
                      </div>
                      <button onClick={()=>setEditActIdx(editActIdx===i?null:i)} style={{padding:"6px 12px",borderRadius:"6px",border:`1px solid ${editActIdx===i?"#f59e0b":"#93c5fd"}`,background:editActIdx===i?"#fffbeb":"#eff6ff",color:editActIdx===i?"#92400e":"#1d4ed8",cursor:"pointer",fontSize:"12px",fontWeight:"600",flexShrink:0}}>{editActIdx===i?"▲ Close":"✏️ Edit"}</button>
                      <button onClick={()=>{setActs(p=>p.filter(x=>x._k!==a._k));if(editActIdx===i)setEditActIdx(null);}} style={{padding:"7px 12px",borderRadius:"7px",border:"1px solid #fca5a5",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"13px",flexShrink:0}}>✕</button>
                    </div>
                  ))}
                </Card>
              )}
              {editActIdx!==null&&acts[editActIdx]&&(
                <Card style={{border:`2px solid ${AM}`,borderTop:`4px solid ${AM}`,marginBottom:"14px"}}>
                  <div style={{fontWeight:"700",fontSize:"15px",color:"#92400e",marginBottom:"14px"}}>✏️ Editing Activity #{editActIdx+1}</div>
                  <ActivityNew key={"edit-"+editActIdx} initialData={acts[editActIdx]} onAdd={a=>{setActs(p=>p.map((x,idx)=>idx===editActIdx?{...a,_k:x._k}:x));setEditActIdx(null);flash("✓ Activity updated");}} wt={workTypes} editMode lists={lists}/>
                </Card>
              )}
              <Card style={{border:"2px solid "+AM,borderTop:"4px solid "+AM}}>
                <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"18px",paddingBottom:"12px",borderBottom:"1px solid #e5e7eb"}}>
                  <div style={{background:AM,padding:"10px",borderRadius:"10px",flexShrink:0,display:"flex"}}><i className="ti ti-tools" style={{fontSize:"20px",color:"#fff",display:"block"}} aria-hidden/></div>
                  <div><div style={{fontWeight:"800",fontSize:"16px",color:"#92400e"}}>Add a Work Activity</div><div style={{fontSize:"13px",color:"#6b7280"}}>Fill all sections, then tap "Add to My List". Repeat for each activity.</div></div>
                </div>
              {acts.length>0&&<div style={{borderTop:"2px dashed #d1fae5",margin:"4px 0 16px",textAlign:"center"}}><span style={{background:"#fff",padding:"0 10px",fontSize:"11px",color:"#6b7280",fontWeight:"600",position:"relative",top:"-10px"}}>+ Add Another Activity</span></div>}
                <ActivityNew onAdd={a=>{setActs(p=>[...p,a]);setHasUnsavedForm(false);flash("✓ Work activity added");}} wt={workTypes} lists={lists} onDirty={()=>setHasUnsavedForm(true)} requiredFields={globalLists.requiredFields}/>
              </Card>
            </div>
          )}

          {/* Step 4: Review */}
          {step===3&&(
            <Card style={{borderTop:"4px solid "+GN}}>
              <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"18px",paddingBottom:"12px",borderBottom:"1px solid #e5e7eb"}}>
                <div style={{background:GN,padding:"10px",borderRadius:"10px",flexShrink:0,display:"flex"}}><i className="ti ti-send" style={{fontSize:"22px",color:"#fff",display:"block"}} aria-hidden/></div>
                <div><div style={{fontWeight:"800",fontSize:"17px",color:GN}}>Review & Submit</div><div style={{fontSize:"13px",color:"#6b7280"}}>Read carefully before submitting.</div></div>
              </div>
              {/* Warnings */}
              {(()=>{
                const ws=[];
                if(acts.length===0&&matTxs.length===0)ws.push({icon:"❌",msg:"Nothing added — go back and add work activities or material moves",c:RD,bg:"#fef2f2",bc:"#fca5a5"});
                if(acts.length>0&&acts.some(a=>!a.theoQty&&!a.prodQty))ws.push({icon:"⚠️",msg:acts.filter(a=>!a.theoQty&&!a.prodQty).length+" activit"+(acts.filter(a=>!a.theoQty&&!a.prodQty).length===1?"y has":"ies have")+" no quantity filled",c:"#92400e",bg:"#fffbeb",bc:"#f59e0b"});
                if(!hdr.incharge)ws.push({icon:"⚠️",msg:"No Incharge — this DPR won't appear in anyone's approval list",c:"#92400e",bg:"#fffbeb",bc:"#f59e0b"});
                const daysAgo=Math.floor((new Date()-new Date(hdr.date+"T12:00:00"))/86400000);
                if(daysAgo>2)ws.push({icon:"🔒",msg:"Date is "+daysAgo+" days in the past",c:"#92400e",bg:"#fffbeb",bc:"#f59e0b"});
                if(ws.length===0)return<div style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:"8px",padding:"10px 14px",marginBottom:"12px",fontSize:"13px",fontWeight:"600",color:"#166534"}}>✅ Everything looks good — ready to submit!</div>;
                return ws.map((w,i)=><div key={i} style={{background:w.bg,border:`1px solid ${w.bc}`,borderRadius:"8px",padding:"10px 14px",marginBottom:"8px",fontSize:"13px",fontWeight:"600",color:w.c,display:"flex",alignItems:"center",gap:"8px"}}>{w.icon} {w.msg}</div>);
              })()}
              <div style={{background:"#f8fafc",border:"1px solid #e5e7eb",borderRadius:"10px",padding:"14px",marginBottom:"14px"}}>
                <Grid cols={mobile?"1fr 1fr":"repeat(5,1fr)"}>
                  {[["Date",hdr.date],["Engineer",hdr.engineer==="__other"?(hdr.engCustom||"—"):hdr.engineer||"—"],["Shift",hdr.shift],["Dept",(hdr.dept||"—").split(" ")[0]],["Incharge",hdr.incharge||"—"]].map(([k,v])=>(
                    <div key={k}><div style={{color:"#9ca3af",fontSize:"11px",fontWeight:"600",marginBottom:"3px"}}>{k}</div><div style={{fontWeight:"700",color:v==="—"?"#dc2626":"#111827",fontSize:"13px"}}>{v}</div></div>
                  ))}
                </Grid>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px"}}>
                <div style={{background:"#faf5ff",border:`2px solid ${acts.length===0?"#dc2626":"#c4b5fd"}`,borderRadius:"10px",padding:"14px",textAlign:"center"}}>
                  <div style={{fontSize:"36px",fontWeight:"800",color:acts.length===0?RD:PU}}>{acts.length}</div>
                  <div style={{fontSize:"13px",color:acts.length===0?RD:"#7c3aed",fontWeight:"700"}}>Work Activities</div>
                </div>
                <div style={{background:"#f0fdfa",border:"2px solid #5eead4",borderRadius:"10px",padding:"14px",textAlign:"center"}}>
                  <div style={{fontSize:"36px",fontWeight:"800",color:"#0f766e"}}>{matTxs.length}</div>
                  <div style={{fontSize:"13px",color:"#0f766e",fontWeight:"700"}}>Material Movements</div>
                </div>
              </div>
              {/* DPR edit history */}
              {(()=>{
                const existing=editingSubId?subs.find(s=>s.id===editingSubId):null;
                const hist=existing?.editHistory||[];
                if(hist.length===0)return null;
                return(<div style={{marginTop:"12px",background:"#f8fafc",borderRadius:"8px",padding:"10px 12px"}}>
                  <div style={{fontWeight:"700",fontSize:"12px",color:"#374151",marginBottom:"7px",textTransform:"uppercase",letterSpacing:"0.04em"}}>📝 Edit History ({hist.length} revision{hist.length!==1?"s":""})</div>
                  {hist.map((h,i)=><div key={i} style={{fontSize:"12px",color:"#6b7280",padding:"4px 0",borderBottom:"1px solid #f0f0f0"}}>
                    {new Date(h.editedAt).toLocaleDateString("en-IN",{day:"2-digit",month:"short"})} at {new Date(h.editedAt).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"})} — edited by <strong>{h.editedBy}</strong> · {h.prevActCount}→{h.newActCount} activities
                  </div>)}
                </div>);
              })()}
            </Card>
          )}

          {/* Navigation buttons — sticky above the tab bar on mobile so Next/Submit is always reachable */}
          <div style={{display:"flex",justifyContent:"space-between",marginTop:"16px",gap:"10px",...(mobile?{position:"sticky",bottom:"64px",zIndex:500,background:"#f8fafc",padding:"10px 2px calc(10px + env(safe-area-inset-bottom))",margin:"16px -2px 0",borderTop:"1px solid #e5e7eb",boxShadow:"0 -6px 8px -8px rgba(0,0,0,.25)"}:{})}}>
            <button onClick={()=>{if(step>0){setStep(s=>s-1);}else{safeNav(()=>setView("dashboard"));}}} style={{padding:mobile?"15px 18px":"13px 20px",borderRadius:"10px",border:"2px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:mobile?"15px":"14px",fontWeight:"700",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px",color:"#374151",flex:mobile?"1":"0 0 auto"}}>
              <i className="ti ti-arrow-left" aria-hidden/>{step===0?"Dashboard":"Previous"}
            </button>
            <div style={{display:"flex",gap:"8px",flex:mobile?"1":"0 0 auto"}}>
              {step<3&&<button onClick={()=>{
                // Validate Step 1 required fields
                if(step===0){
                  const eng=roleKey(user?.role)==="engineer"?user.name:(hdr.engineer==="__other"?(hdr.engCustom||"").trim():hdr.engineer);
                  const errs={};
                  if(!eng)errs.engineer="Required";
                  if(!hdr.date)errs.date="Required";
                  if(!hdr.dept)errs.dept="Required";
                  // Admin-configured required fields for header
                  const rf=globalLists.requiredFields||{};
                  if(rf.incharge&&!hdr.incharge)errs.incharge="Required by admin";
                  if(rf.weather&&!hdr.weather)errs.weather="Required by admin";
                  if(rf.difficulty&&!hdr.difficulty)errs.difficulty="Required by admin";
                  if(Object.keys(errs).length>0){setStepErrors(errs);flash("Please fill required fields","err");return;}
                  setStepErrors({});
                }
                // Warn about unsaved form data (activity or material)
                if(hasUnsavedForm){
                  if(!window.confirm("⚠️ You have filled data that hasn't been added to the list yet.\n\nClick the Add button first to save it, or click OK to discard and move on."))return;
                  setHasUnsavedForm(false);
                }
                setStep(s=>s+1);
              }} style={{flex:mobile?1:undefined,padding:mobile?"15px 22px":"13px 24px",borderRadius:"10px",border:"none",background:AM,color:"#fff",cursor:"pointer",fontSize:mobile?"15px":"14px",fontWeight:"800",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px"}}>Next <i className="ti ti-arrow-right" aria-hidden/></button>}
              {step===3&&<button onClick={submitDPR} style={{flex:mobile?1:undefined,padding:mobile?"15px 24px":"13px 28px",borderRadius:"10px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:mobile?"16px":"15px",fontWeight:"800",display:"flex",alignItems:"center",justifyContent:"center",gap:"7px"}}><i className="ti ti-send" aria-hidden/>Submit DPR</button>}
            </div>
          </div>
          </>}
        </div>
      )}

      {/* ═══ APPROVE ═══ */}
      {view==="approve"&&caps.approve&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"14px",flexWrap:"wrap"}}>
            <span style={{fontWeight:"700",color:"#374151",fontSize:"14px"}}>Approve DPR for:</span>
            <input type="date" value={apvDate} onChange={e=>setApvDate(e.target.value)} style={{fontSize:"14px",padding:"8px 10px",borderRadius:"8px",border:"1.5px solid #d1d5db"}}/>
            <button onClick={()=>setApvDate(new Date().toISOString().slice(0,10))} style={{fontSize:"13px",padding:"8px 14px",borderRadius:"8px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontWeight:"700"}}>Today</button>
            <button onClick={()=>setApvDate(new Date(Date.now()-86400000).toISOString().slice(0,10))} style={{fontSize:"13px",padding:"8px 14px",borderRadius:"8px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",color:"#6b7280",fontWeight:"600"}}>Yesterday</button>
          </div>
          {/* Pending dates calendar */}
          {(()=>{
            const pendingSubs=subs.filter(s=>canApprove(s)&&!s.approved&&!s.needsRevision);
            const pendingDates=[...new Set(pendingSubs.map(s=>s.date))].sort();
            if(pendingDates.length===0)return null;
            return(
              <div style={{marginBottom:"14px",background:"#fffbeb",border:"2px solid #f59e0b",borderRadius:"12px",padding:"12px 16px"}}>
                <div style={{fontWeight:"700",fontSize:"13px",color:"#92400e",marginBottom:"10px"}}>
                  📅 Dates with pending approvals ({pendingDates.length} date{pendingDates.length!==1?"s":""}, {pendingSubs.length} DPR{pendingSubs.length!==1?"s":""})
                </div>
                <div style={{display:"flex",gap:"6px",flexWrap:"wrap"}}>
                  {pendingDates.map(d=>{
                    const count=pendingSubs.filter(s=>s.date===d).length;
                    const isSelected=apvDate===d;
                    return(
                      <button key={d} onClick={()=>setApvDate(d)}
                        style={{padding:"6px 12px",borderRadius:"8px",border:`2px solid ${isSelected?"#d97706":"#fde68a"}`,background:isSelected?"#f59e0b":"#fff",cursor:"pointer",fontSize:"12px",fontWeight:"700",color:isSelected?"#fff":"#92400e",display:"flex",flexDirection:"column",alignItems:"center",gap:"2px"}}>
                        <span>{new Date(d+"T12:00:00").toLocaleDateString("en-IN",{day:"2-digit",month:"short"})}</span>
                        <span style={{fontSize:"10px",background:isSelected?"rgba(255,255,255,.3)":"#fef3c7",borderRadius:"4px",padding:"1px 5px"}}>{count} pending</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            );
          })()}
          {user&&<div style={{background:"#f0fdf4",border:"2px solid #86efac",borderRadius:"12px",padding:"12px 16px",marginBottom:"14px",display:"flex",alignItems:"center",gap:"12px"}}>
            <Av name={user.name} sz={40}/>
            <div><div style={{fontWeight:"700",fontSize:"14px",color:GN}}>Approving as: {user.name}</div><div style={{marginTop:"3px",display:"flex",gap:"6px",alignItems:"center"}}><RoleB role={user.role}/>{roleKey(user.role)==="incharge"&&<span style={{fontSize:"12px",color:"#6b7280"}}>You can only approve your engineers.</span>}</div></div>
          </div>}
          {apvS.filter(s=>canApprove(s)).length===0
            ?<Card style={{textAlign:"center",padding:"3rem",color:"#9ca3af"}}><i className="ti ti-inbox" style={{fontSize:"32px",display:"block",marginBottom:"8px"}} aria-hidden/>No submissions to approve for this date.</Card>
            :<div style={{display:"flex",flexDirection:"column",gap:"12px"}}>
              {apvS.filter(s=>canApprove(s)).map(s=>(
                <Card key={s.id} style={{borderLeft:`4px solid ${s.approved?GN:AM}`}}>
                  <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:"10px",marginBottom:"14px",paddingBottom:"12px",borderBottom:"1px solid #e5e7eb",flexWrap:"wrap"}}>
                    <div style={{display:"flex",alignItems:"center",gap:"10px"}}>
                      <Av name={s.engineer} sz={44}/>
                      <div>
                        <div style={{fontWeight:"800",fontSize:"16px",color:"#111827"}}>{s.engineer}</div>
                        <div style={{display:"flex",gap:"5px",alignItems:"center",marginTop:"5px",flexWrap:"wrap"}}>
                          <DeptB dept={s.dept}/>
                          <Pill label={s.shift+" Shift"} color={s.shift==="Day"?"#d97706":"#2563eb"} bg={s.shift==="Day"?"#fffbeb":"#eff6ff"}/>
                          <span style={{fontSize:"12px",color:"#6b7280"}}>{(s.activities||[]).length} activit{(s.activities||[]).length===1?"y":"ies"}</span>
                          {(s.matTxs||[]).length>0&&<span style={{fontSize:"12px",color:"#6b7280"}}>{(s.matTxs||[]).length} mat.moves</span>}
                        </div>
                      </div>
                    </div>
                    {s.approved&&<div style={{textAlign:"right"}}>
                      <Pill label="✓ Approved" color={GN} bg="#f0fdf4"/>
                      <div style={{fontSize:"12px",color:"#6b7280",marginTop:"4px",fontWeight:"600"}}>by {s.approvedBy}</div>
                      {s.approvalNote&&<div style={{fontSize:"11px",color:"#9ca3af",fontStyle:"italic"}}>"{s.approvalNote}"</div>}
                      {s.editHistory&&s.editHistory.length>0&&(
                        <div style={{marginTop:"6px",padding:"6px 10px",background:"#f8fafc",borderRadius:"6px",fontSize:"11px",color:"#6b7280"}}>
                          📝 {s.editHistory.length} edit{s.editHistory.length!==1?"s":""} — last by {s.editHistory[s.editHistory.length-1].editedBy} on {new Date(s.editHistory[s.editHistory.length-1].editedAt).toLocaleDateString("en-IN",{day:"2-digit",month:"short"})}
                        </div>
                      )}
                    </div>}
                    {!s.approved&&<Pill label="⏳ Pending" color="#d97706" bg="#fffbeb"/>}
                  </div>
                  {(s.activities||[]).map((a,i)=>(
                    <div key={a._k||i} style={{padding:"10px 14px",borderRadius:"8px",background:"#f8fafc",fontSize:"13px",border:"1px solid #f3f4f6",marginBottom:"6px"}}>
                      <span style={{fontWeight:"700"}}>{i+1}. {a.actType||(a.actCustom||"—")}</span>{a.desc&&<span style={{color:"#6b7280"}}> — {a.desc}</span>}
                      <div style={{fontFamily:"monospace",color:"#9ca3af",fontSize:"11px",marginTop:"3px"}}>{[a.chFrom&&`CH ${a.chFrom}→${a.chTo||"?"}`,a.side&&a.side,a.theoQty&&`${a.theoQty} ${a.unit}`,(a.assets||[]).length>0&&`${a.assets.length} machine(s)`,(a.contractors||[]).length>0&&a.contractors.map(c=>c.name).join(", ")].filter(Boolean).join(" · ")}</div>
                    </div>
                  ))}
                  {(s.matTxs||[]).map((m,i)=>(
                    <div key={m._k||i} style={{padding:"10px 14px",borderRadius:"8px",background:"#f0fdfa",fontSize:"13px",border:"1px solid #ccfbf1",marginBottom:"6px"}}>
                      <span style={{fontWeight:"700",color:"#0f766e"}}>📦 {m.material==="Other"?m.matCustom:m.material}</span>
                      <Pill label={m.txType} color="#0f766e" bg="#f0fdfa"/>
                      {m.recvCH&&<span style={{color:"#6b7280",fontSize:"11px"}}> | Recv:{m.recvCH}</span>}
                      {m.sendFromCH&&<span style={{color:"#6b7280",fontSize:"11px"}}> | Send:{m.sendFromCH}→{m.sendToCH}</span>}
                    </div>
                  ))}
                  {apvId===s.id?(
                    <div style={{background:"#f8fafc",border:"1px solid #e5e7eb",borderRadius:"10px",padding:"14px",marginTop:"10px"}}>
                      <div style={{marginBottom:"12px"}}><L t="Comment (optional)"/><Inp value={apvNote} onChange={e=>setApvNote(e.target.value)} placeholder="Any observations or instructions…"/></div>
                      <div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>
                        <button onClick={()=>approveSub(s.id)} style={{flex:1,padding:"13px",borderRadius:"10px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:"15px",fontWeight:"800",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px"}}><i className="ti ti-check" aria-hidden/>Approve</button>
                        <button onClick={()=>rejectSub(s.id)} style={{flex:1,padding:"13px",borderRadius:"10px",border:"2px solid #fca5a5",background:"#fef2f2",color:RD,cursor:"pointer",fontSize:"14px",fontWeight:"700"}}>↩ Send for Revision</button>
                        {user?.role==='admin'&&<button onClick={()=>deleteSub(s.id)} title="Admin: Delete this DPR" style={{padding:"13px 14px",borderRadius:"10px",border:"2px solid #dc2626",background:"#dc2626",color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700",flexShrink:0}}>🗑 Delete</button>}
                        <button onClick={()=>{setApvId(null);setApvNote("");}} style={{padding:"13px 16px",borderRadius:"10px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"13px",color:"#6b7280"}}>✕</button>
                      </div>
                    </div>
                  ):(<div style={{display:"flex",gap:"8px",marginTop:"10px",flexWrap:"wrap"}}>
                    {s.needsRevision&&<div style={{width:"100%",marginBottom:"6px",padding:"8px 12px",borderRadius:"8px",background:"#fef3c7",border:"1px solid #f59e0b",fontSize:"12px",color:"#92400e",fontWeight:"600"}}>
                      ↩ Sent for revision — note: "{s.revisionNote||"Please check and re-submit"}"
                      <button onClick={()=>{if(!window.confirm("Cancel the revision request? The DPR will go back to Pending status."))return;rtdbPatch(pb()+'/submissions/'+s.id,{needsRevision:false,revisionNote:"",approved:false,approvedBy:"",approvedAt:"",approvalNote:""}).then(()=>flash("✅ Revision cancelled — DPR is now Pending")).catch(e=>flash(e.message,"err"));}} style={{marginLeft:"10px",padding:"3px 10px",borderRadius:"6px",border:"1px solid #d97706",background:"#fff",color:"#92400e",cursor:"pointer",fontSize:"12px",fontWeight:"700"}}>✕ Cancel Revision</button>
                    </div>}
                    <button onClick={()=>setApvId(s.id)} style={{flex:1,padding:"13px",borderRadius:"10px",border:"none",background:NV,color:"#fff",cursor:"pointer",fontSize:"14px",fontWeight:"800",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px"}}><i className="ti ti-circle-check" aria-hidden/>{s.approved?"Revise Approval":"Approve This DPR"}</button>
                    {user?.role==='admin'&&<button onClick={()=>deleteSub(s.id)} title="Admin: Permanently delete this DPR" style={{padding:"13px 16px",borderRadius:"10px",border:"2px solid #dc2626",background:"#fff",color:"#dc2626",cursor:"pointer",fontSize:"13px",fontWeight:"700",flexShrink:0}}>🗑</button>}
                  </div>)}
                </Card>
              ))}
            </div>}
        </div>
      )}

      {/* ═══ MANAGE ENGINEERS ═══ */}
      {view==="manage"&&caps.manage&&(
        <div>
          {/* Engineers from Users who are assigned to this project */}
          {(()=>{
            const projEngUsers=users.filter(u=>roleKey(u.role)==="engineer"&&u.assignedProjectId===activeProject?.id);
            const noIncharge=projEngUsers.filter(u=>{const er=engineers.find(e=>e.name===u.name);return !er?.incharge;});
            return(
              <>
                {noIncharge.length>0&&(
                  <div style={{background:"#fffbeb",border:"2px solid #f59e0b",borderRadius:"10px",padding:"12px 16px",marginBottom:"14px",fontSize:"13px",color:"#92400e",fontWeight:"600"}}>
                    ⚠️ {noIncharge.length} engineer{noIncharge.length>1?"s have":" has"} no incharge assigned: {noIncharge.map(u=>u.name).join(", ")}
                  </div>
                )}
                {projEngUsers.length===0?(
                  <Card style={{textAlign:"center",padding:"3rem",color:"#9ca3af"}}>
                    <div style={{fontSize:"40px",marginBottom:"8px"}}>👷</div>
                    <div style={{fontWeight:"700",fontSize:"15px",marginBottom:"4px"}}>No engineers assigned to this project</div>
                    <div style={{fontSize:"13px",marginBottom:"16px"}}>Go to Global Admin → Users and assign engineers to this project.</div>
                  </Card>
                ):(
                  <Card>
                    <div style={SH}>
                      <span>Project Engineers ({projEngUsers.length})</span>
                      <span style={{fontSize:"11px",color:"#9ca3af",fontWeight:"400",textTransform:"none"}}>Names come from Users tab. Set incharge &amp; designation below.</span>
                    </div>
                    <div style={{overflowX:"auto"}}>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px"}}>
                        <thead><tr style={{background:"#f8fafc"}}>{(mobile?["Name","Incharge","Actions"]:["Name","Dept","Designation","Incharge","Actions"]).map(h=><th key={h} style={{padding:"10px",textAlign:"left",fontWeight:"700",color:"#6b7280",fontSize:"11px",borderBottom:"1px solid #e5e7eb"}}>{h}</th>)}</tr></thead>
                        <tbody>
                          {projEngUsers.map(u=>{
                            // Use user.id as stable key for their engineer record
                            const er=engineers.find(e=>e.name===u.name)||{id:u.id,name:u.name,incharge:"",designation:"",dept:""};
                            return(
                              <EngineerRow key={u.id} e={{...er,id:u.id}}
                                onSave={(id,data)=>{
                                  rtdbPut(pb()+'/engineers/'+u.id,{...er,...data,id:u.id,name:u.name})
                                    .then(()=>flash("✅ Saved")).catch(e=>flash(e.message,"err"));
                                }}
                                onRemove={()=>flash("To remove: go to Users tab and change their project assignment","err")}
                                inchargeOpts={users.filter(x=>roleKey(x.role)==="incharge"&&x.assignedProjectId===activeProject?.id).map(x=>x.name)}
                                designationOpts={lists.designations||DESIGNATIONS}
                                deptOpts={lists.depts||DEPTS}
                              />
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </Card>
                )}
              </>
            );
          })()}
        </div>
      )}

      {/* ═══ USERS & ACCESS ═══ */}
      {/* ═══ MONTHLY REPORT ═══ */}
      {view==="monthly"&&(
        <div>
          {/* Role scope banner */}
          {roleKey(user?.role)==="engineer"&&<div style={{background:"#eff6ff",border:"1px solid #93c5fd",borderRadius:"10px",padding:"10px 16px",marginBottom:"12px",fontSize:"13px",color:"#1e40af",fontWeight:"600",display:"flex",alignItems:"center",gap:"8px"}}><i className="ti ti-user" style={{fontSize:"15px"}} aria-hidden/>Showing your personal DPR data only</div>}
          {roleKey(user?.role)==="incharge"&&(()=>{const myEngs=engineers.filter(e=>e.incharge===user.name);return<div style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:"10px",padding:"10px 16px",marginBottom:"12px",fontSize:"13px",color:"#166534",fontWeight:"600",display:"flex",alignItems:"center",gap:"8px"}}><i className="ti ti-users-group" style={{fontSize:"15px"}} aria-hidden/>Showing data for your engineers: {myEngs.length>0?myEngs.map(e=>e.name).join(", "):"(no engineers assigned yet)"}</div>;})()}
          <Card style={{marginBottom:"14px"}}>
            <div style={{fontWeight:"700",fontSize:"16px",color:NV,marginBottom:"14px"}}>📊 Consolidated DPR Report</div>
            <Grid cols={mobile?"1fr":"1fr 1fr 1fr auto"}>
              <F lbl="From Date"><input type="date" value={reportFrom} onChange={e=>setReportFrom(e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px"}}/></F>
              <F lbl="To Date"><input type="date" value={reportTo} onChange={e=>setReportTo(e.target.value)} style={{width:"100%",boxSizing:"border-box",padding:"11px 12px",borderRadius:"8px",border:"1.5px solid #d1d5db",fontSize:"15px"}}/></F>
              <F lbl="Quick Select">
                <Sel val="" onChange={v=>{const now=new Date();if(v==="allData"){setReportFrom("2020-01-01");setReportTo(now.toISOString().slice(0,10));}else if(v==="thisMonth"){setReportFrom(now.toISOString().slice(0,7)+"-01");setReportTo(now.toISOString().slice(0,10));}else if(v==="lastMonth"){const lm=new Date(now.getFullYear(),now.getMonth()-1,1);const le=new Date(now.getFullYear(),now.getMonth(),0);setReportFrom(lm.toISOString().slice(0,10));setReportTo(le.toISOString().slice(0,10));}else if(v==="last7"){const d=new Date(now-6*86400000);setReportFrom(d.toISOString().slice(0,10));setReportTo(now.toISOString().slice(0,10));}else if(v==="last30"){const d=new Date(now-29*86400000);setReportFrom(d.toISOString().slice(0,10));setReportTo(now.toISOString().slice(0,10));}}} opts={[]} placeholder="— Quick Range —">
                  <option value="thisMonth">This Month</option>
                  <option value="lastMonth">Last Month</option>
                  <option value="last7">Last 7 Days</option>
                  <option value="last30">Last 30 Days</option>
                </Sel>
              </F>
              <F lbl=" ">
                <div style={{display:"flex",gap:"8px",paddingTop:"2px"}}>
                  {caps.download&&<button onClick={()=>{const ms=getReportSubs();if(!ms.length){flash("No data for this range","err");return;}doExcel(ms,engineers,reportFrom+' to '+reportTo,activeProject?.name);}} style={{padding:"11px 16px",borderRadius:"8px",border:"none",background:GN,color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700",display:"flex",alignItems:"center",gap:"6px",whiteSpace:"nowrap"}}><i className="ti ti-file-spreadsheet" aria-hidden/>Download</button>}
                  <button onClick={()=>setShareOpen(true)} style={{padding:"11px 14px",borderRadius:"8px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"700",display:"flex",alignItems:"center",gap:"6px",whiteSpace:"nowrap"}}><i className="ti ti-share" aria-hidden/>Share</button>
                </div>
              </F>
            </Grid>
          </Card>
          {(()=>{
            const ms=getReportSubs();
            const days=[...new Set(ms.map(s=>s.date))].sort();
            if(ms.length===0)return<Card style={{textAlign:"center",padding:"3rem",color:"#9ca3af"}}><i className="ti ti-calendar-off" style={{fontSize:"32px",display:"block",marginBottom:"8px"}} aria-hidden/>No data for {reportFrom} → {reportTo}</Card>;
            return(
              <>
                <div style={{display:"grid",gridTemplateColumns:mobile?"repeat(2,1fr)":"repeat(4,1fr)",gap:"10px",marginBottom:"16px"}}>
                  {[{l:"Working Days",v:days.length,c:NV,bg:"#eff6ff"},{l:"Total Submissions",v:ms.length,c:PU,bg:"#f5f3ff"},{l:"Approved",v:ms.filter(s=>s.approved).length,c:GN,bg:"#f0fdf4"},{l:"Pending",v:ms.filter(s=>!s.approved).length,c:"#d97706",bg:"#fffbeb"}].map(m=>(
                    <div key={m.l} style={{background:m.bg,borderRadius:"12px",padding:"14px",border:`1px solid ${m.c}20`}}>
                      <div style={{fontSize:"11px",color:"#6b7280",fontWeight:"600",marginBottom:"6px"}}>{m.l}</div>
                      <div style={{fontSize:"28px",fontWeight:"800",color:m.c}}>{m.v}</div>
                    </div>
                  ))}
                </div>
                <Card style={{marginBottom:"14px"}}>
                  <div style={SH}><span>Engineer Summary — {reportFrom} to {reportTo}</span></div>
                  <div style={{overflowX:"auto"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:"13px"}}>
                      <thead><tr style={{background:"#f8fafc"}}>{["Engineer","Dept","Days Present","Submissions","Activities","Mat.Moves","Approved","Pending"].map(h=><th key={h} style={{padding:"9px 10px",textAlign:"left",fontWeight:"700",color:"#6b7280",fontSize:"11px",borderBottom:"1px solid #e5e7eb",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                      <tbody>{engineers.map(e=>{
                        const es=ms.filter(s=>s.engineer===e.name);
                        if(es.length===0)return null;
                        const daysPresent=[...new Set(es.map(s=>s.date))].length;
                        const acts=es.reduce((a,s)=>a+(s.activities||[]).length,0);
                        const mats=es.reduce((a,s)=>a+(s.matTxs||[]).length,0);
                        const apvd=es.filter(s=>s.approved).length;
                        return(
                          <tr key={e.id} style={{borderBottom:"1px solid #f3f4f6"}}>
                            <td style={{padding:"10px",fontWeight:"700"}}>{e.name}</td>
                            <td style={{padding:"10px"}}><DeptB dept={e.dept}/></td>
                            <td style={{padding:"10px",fontWeight:"700",color:NV}}>{daysPresent}</td>
                            <td style={{padding:"10px",fontWeight:"700",color:PU}}>{es.length}{es.length>daysPresent&&<span style={{background:"#818cf8",color:"#fff",borderRadius:"8px",padding:"1px 6px",fontSize:"10px",marginLeft:"6px"}}>multi</span>}</td>
                            <td style={{padding:"10px",fontWeight:"700",color:acts>0?PU:"#9ca3af"}}>{acts||"—"}</td>
                            <td style={{padding:"10px",fontWeight:"700",color:mats>0?"#0f766e":"#9ca3af"}}>{mats||"—"}</td>
                            <td style={{padding:"10px"}}><Pill label={apvd} color={GN} bg="#f0fdf4"/></td>
                            <td style={{padding:"10px"}}><Pill label={es.length-apvd} color={es.length-apvd>0?"#d97706":"#9ca3af"} bg={es.length-apvd>0?"#fffbeb":"#f3f4f6"}/></td>
                          </tr>
                        );
                      })}</tbody>
                    </table>
                  </div>
                </Card>
                <Card>
                  <div style={SH}><span>Day-wise Attendance — {reportFrom} to {reportTo}</span></div>
                  <div style={{overflowX:"auto"}}>
                    <table style={{borderCollapse:"collapse",fontSize:"12px",minWidth:"500px",width:"100%"}}>
                      <thead><tr style={{background:"#f8fafc"}}>{["Date","Day","Present","Absent","Activities","Approved"].map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",fontWeight:"700",color:"#6b7280",borderBottom:"1px solid #e5e7eb",whiteSpace:"nowrap",fontSize:"11px"}}>{h}</th>)}</tr></thead>
                      <tbody>{days.map(d=>{
                        const ds=ms.filter(s=>s.date===d);
                        const scopedEngCount=roleKey(user?.role)==="engineer"?1:roleKey(user?.role)==="incharge"?engineers.filter(e=>e.incharge===user.name).length:engineers.length;
                        const pres=[...new Set(ds.map(s=>s.engineer))].length;
                        const acts=ds.reduce((a,s)=>a+(s.activities||[]).length,0);
                        const apvd=ds.filter(s=>s.approved).length;
                        const dt=new Date(d+"T12:00:00");
                        return(
                          <tr key={d} style={{borderBottom:"1px solid #f3f4f6"}}>
                            <td style={{padding:"8px 10px",fontWeight:"700",fontFamily:"monospace"}}>{d}</td>
                            <td style={{padding:"8px 10px",color:"#6b7280"}}>{dt.toLocaleDateString("en-IN",{weekday:"short"})}</td>
                            <td style={{padding:"8px 10px"}}><Pill label={"✓ "+pres} color={GN} bg="#f0fdf4"/></td>
                            <td style={{padding:"8px 10px"}}><Pill label={scopedEngCount-pres} color={scopedEngCount-pres>0?RD:"#9ca3af"} bg={scopedEngCount-pres>0?"#fef2f2":"#f3f4f6"}/></td>
                            <td style={{padding:"8px 10px",fontWeight:"700",color:acts>0?PU:"#9ca3af"}}>{acts||"—"}</td>
                            <td style={{padding:"8px 10px"}}><span style={{color:apvd===ds.length?GN:"#d97706",fontWeight:"700"}}>{apvd}/{ds.length}</span></td>
                          </tr>
                        );
                      })}</tbody>
                    </table>
                  </div>
                </Card>
                {/* Chainage Progress Tracker — management/admin only */}
                {(roleKey(user?.role)==="admin"||roleKey(user?.role)==="management")&&(()=>{
                  const workMap={};
                  ms.forEach(s=>(s.activities||[]).forEach(a=>{
                    if(!a.chFrom||!a.chTo)return;
                    const k=a.actType||(a.actCustom||"Work");
                    if(!workMap[k])workMap[k]=[];
                    workMap[k].push({from:a.chFrom,to:a.chTo,eng:s.engineer,date:s.date,qty:a.theoQty,unit:a.unit});
                  }));
                  if(Object.keys(workMap).length===0)return null;
                  function chToKm(ch){const[km,m]=(ch||"0+0").split("+").map(Number);return(km||0)+(m||0)/1000;}
                  return(
                    <Card style={{marginBottom:"14px"}}>
                      <div style={SH}><span>📍 Chainage Progress — {reportFrom} to {reportTo}</span></div>
                      <div style={{display:"flex",flexDirection:"column",gap:"12px"}}>
                        {Object.entries(workMap).map(([wt,entries])=>{
                          const allKm=entries.map(e=>({f:chToKm(e.from),t:chToKm(e.to)}));
                          const minKm=Math.min(...allKm.map(x=>x.f));
                          const maxKm=Math.max(...allKm.map(x=>x.t));
                          const total=(maxKm-minKm).toFixed(3);
                          return(
                            <div key={wt} style={{padding:"12px",borderRadius:"8px",background:"#f8fafc",border:"1px solid #e5e7eb"}}>
                              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"8px",flexWrap:"wrap",gap:"6px"}}>
                                <div style={{fontWeight:"700",fontSize:"13px",color:NV}}>{wt}</div>
                                <div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>
                                  <span style={{fontSize:"12px",color:"#6b7280"}}>CH {entries[0]?.from} → {entries[entries.length-1]?.to}</span>
                                  <span style={{fontSize:"12px",fontWeight:"700",color:PU}}>{total} km covered</span>
                                  <span style={{fontSize:"12px",color:"#6b7280"}}>{entries.length} session{entries.length!==1?"s":""}</span>
                                </div>
                              </div>
                              <div style={{position:"relative",height:"8px",background:"#e5e7eb",borderRadius:"4px",overflow:"visible"}}>
                                {allKm.map((seg,i)=>{
                                  const range=maxKm-minKm||1;
                                  const left=((seg.f-minKm)/range)*100;
                                  const width=((seg.t-seg.f)/range)*100;
                                  return<div key={i} title={`${entries[i].from} → ${entries[i].to} (${entries[i].eng})`} style={{position:"absolute",left:left+"%",width:Math.max(width,1)+"%",height:"100%",background:NV,borderRadius:"3px",opacity:0.7+0.3*(i/allKm.length)}}/>;
                                })}
                              </div>
                              <div style={{display:"flex",justifyContent:"space-between",marginTop:"4px",fontSize:"10px",color:"#9ca3af"}}>
                                <span>CH {entries[0]?.from}</span><span>CH {entries[entries.length-1]?.to}</span>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </Card>
                  );
                })()}
              </>
            );
          })()}
        </div>
      )}

      {/* ═══ SETTINGS ═══ */}
      {view==="settings"&&caps.settings&&(
        <SettingsPanel
          settingsTab={settingsTab} setSettingsTab={setSettingsTab}
          listEdits={listEdits} setListEdits={setListEdits}
          lists={lists} setLists={setLists}
          workTypes={workTypes} saveWt={saveWt}
          flash={flash} projectId={activeProject?.id}
          globalLists={globalLists}
        />
      )}

      </div>

      {/* ═══ SHARE MODAL ═══ */}
      
      {shareOpen&&<ShareModal subs={roleKey(user?.role)==="engineer"||roleKey(user?.role)==="incharge"?getReportSubs():subs} engineers={engineers} dashDate={dashDate} reportFrom={reportFrom} reportTo={reportTo} onClose={()=>setShareOpen(false)} flash={flash}/>}
    </div>
  );
}
