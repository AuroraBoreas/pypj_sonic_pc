from single import single_meas

if __name__ == "__main__":
    #multiple times
    meas_N_times = 10
    i = 0
    while i < meas_N_times:
        single_meas()
        i += 1